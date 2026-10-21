//! Running external programs for Lua (`xl.run`, `xl.shell`): captured output, a timeout that
//! really kills the process, no console window flashing up inside Excel.

use std::io::{Read, Write};
use std::path::PathBuf;
use std::process::{Command, Stdio};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, mpsc};
use std::time::{Duration, Instant};

/// Output beyond this per stream is dropped (and reported), so a runaway program can't eat memory.
pub const MAX_OUTPUT: usize = 32 * 1024 * 1024;
pub const DEFAULT_TIMEOUT: Duration = Duration::from_secs(60);
pub const MAX_TIMEOUT: Duration = Duration::from_secs(3600);
/// How long to wait for the output readers after killing a process whose children may still hold
/// the pipes open.
const READER_GRACE: Duration = Duration::from_secs(2);

#[derive(Clone, Debug)]
pub struct RunSpec {
    pub program: String,
    pub args: Vec<String>,
    /// Windows only: the command tail passed verbatim (no re-quoting), used for `cmd /S /C "…"`.
    pub raw_tail: Option<String>,
    pub cwd: Option<PathBuf>,
    pub stdin: Option<Vec<u8>>,
    pub timeout: Duration,
    /// When set from another thread, the process is killed and the run reports `cancelled`.
    pub cancel: Option<Arc<AtomicBool>>,
}

impl RunSpec {
    pub fn new(program: impl Into<String>, args: Vec<String>) -> Self {
        RunSpec {
            program: program.into(),
            args,
            raw_tail: None,
            cwd: None,
            stdin: None,
            timeout: DEFAULT_TIMEOUT,
            cancel: None,
        }
    }

    /// Runs `line` through the platform shell (`cmd /S /C` or `sh -c`).
    pub fn shell(line: &str) -> Self {
        if cfg!(windows) {
            let mut s = RunSpec::new("cmd", Vec::new());
            s.raw_tail = Some(format!("/S /C \"{line}\""));
            s
        } else {
            RunSpec::new("sh", vec!["-c".to_owned(), line.to_owned()])
        }
    }
}

#[derive(Debug, Default)]
pub struct RunResult {
    /// `None` when the process was killed (timeout) or ended without an exit code.
    pub code: Option<i32>,
    pub stdout: Vec<u8>,
    pub stderr: Vec<u8>,
    pub timed_out: bool,
    /// Killed because the caller asked to stop.
    pub cancelled: bool,
    pub truncated: bool,
}

impl RunResult {
    pub fn ok(&self) -> bool {
        self.code == Some(0) && !self.timed_out && !self.cancelled
    }
}

fn read_capped<R: Read + Send + 'static>(mut r: R) -> mpsc::Receiver<(Vec<u8>, bool)> {
    let (tx, rx) = mpsc::channel();
    std::thread::spawn(move || {
        let mut out = Vec::new();
        let mut buf = [0u8; 16 * 1024];
        let mut truncated = false;
        loop {
            match r.read(&mut buf) {
                Ok(0) | Err(_) => break,
                Ok(n) => {
                    let room = MAX_OUTPUT.saturating_sub(out.len());
                    truncated |= n > room;
                    out.extend_from_slice(&buf[..n.min(room)]);
                }
            }
        }
        let _ = tx.send((out, truncated));
    });
    rx
}

#[cfg(windows)]
fn kill_tree(pid: u32) {
    use std::os::windows::process::CommandExt;
    // Best effort: also end the program's children (cmd -> ping, ...).
    let _ = Command::new("taskkill")
        .args(["/T", "/F", "/PID", &pid.to_string()])
        .creation_flags(0x0800_0000)
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status();
}

#[cfg(not(windows))]
fn kill_tree(_pid: u32) {}

pub fn run(spec: &RunSpec) -> std::io::Result<RunResult> {
    let mut cmd = Command::new(&spec.program);
    cmd.args(&spec.args);
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        if let Some(tail) = &spec.raw_tail {
            cmd.raw_arg(tail);
        }
        cmd.creation_flags(0x0800_0000); // CREATE_NO_WINDOW: no console flashing over Excel
    }
    if let Some(dir) = &spec.cwd {
        cmd.current_dir(dir);
    }
    cmd.stdin(if spec.stdin.is_some() {
        Stdio::piped()
    } else {
        Stdio::null()
    })
    .stdout(Stdio::piped())
    .stderr(Stdio::piped());
    let mut child = cmd.spawn()?;

    if let (Some(data), Some(mut pipe)) = (spec.stdin.clone(), child.stdin.take()) {
        std::thread::spawn(move || {
            let _ = pipe.write_all(&data); // a program that exits early closes the pipe: fine
        });
    }
    let out_rx = read_capped(child.stdout.take().expect("piped"));
    let err_rx = read_capped(child.stderr.take().expect("piped"));

    let timeout = spec.timeout.min(MAX_TIMEOUT);
    let start = Instant::now();
    let mut result = RunResult::default();
    loop {
        if let Some(status) = child.try_wait()? {
            result.code = status.code();
            break;
        }
        let cancelled = spec
            .cancel
            .as_ref()
            .is_some_and(|c| c.load(Ordering::Relaxed));
        if cancelled || start.elapsed() >= timeout {
            let pid = child.id();
            kill_tree(pid);
            let _ = child.kill();
            let _ = child.wait();
            result.timed_out = !cancelled;
            result.cancelled = cancelled;
            break;
        }
        std::thread::sleep(Duration::from_millis(10));
    }
    let grace = if result.timed_out || result.cancelled {
        READER_GRACE
    } else {
        Duration::from_secs(30)
    };
    if let Ok((o, t)) = out_rx.recv_timeout(grace) {
        result.stdout = o;
        result.truncated |= t;
    }
    if let Ok((e, t)) = err_rx.recv_timeout(grace) {
        result.stderr = e;
        result.truncated |= t;
    }
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn text(b: &[u8]) -> String {
        String::from_utf8_lossy(b).trim().to_owned()
    }

    #[test]
    fn captures_output_and_exit_codes() {
        let r = run(&RunSpec::shell("echo hello")).unwrap();
        assert!(r.ok(), "{r:?}");
        assert_eq!(text(&r.stdout), "hello");

        let r = run(&RunSpec::shell("exit 3")).unwrap();
        assert_eq!(r.code, Some(3));
        assert!(!r.ok() && !r.timed_out);
    }

    #[test]
    fn cancelling_kills_the_program_promptly() {
        let line = if cfg!(windows) {
            "ping -n 30 127.0.0.1 > nul"
        } else {
            "sleep 30"
        };
        let cancel = Arc::new(AtomicBool::new(false));
        let mut spec = RunSpec::shell(line);
        spec.cancel = Some(Arc::clone(&cancel));
        let flag = Arc::clone(&cancel);
        std::thread::spawn(move || {
            std::thread::sleep(Duration::from_millis(200));
            flag.store(true, Ordering::Relaxed);
        });
        let start = Instant::now();
        let r = run(&spec).unwrap();
        assert!(r.cancelled && !r.timed_out && !r.ok(), "{r:?}");
        assert!(start.elapsed() < Duration::from_secs(10));
    }

    #[test]
    fn stderr_is_separate() {
        let line = if cfg!(windows) {
            "echo oops 1>&2"
        } else {
            "echo oops >&2"
        };
        let r = run(&RunSpec::shell(line)).unwrap();
        assert_eq!(
            (text(&r.stdout).as_str(), text(&r.stderr).as_str()),
            ("", "oops")
        );
    }

    #[test]
    fn stdin_reaches_the_program() {
        let mut s = RunSpec::shell(if cfg!(windows) {
            "findstr \"^\""
        } else {
            "cat"
        });
        s.stdin = Some(b"from stdin\n".to_vec());
        let r = run(&s).unwrap();
        assert_eq!(text(&r.stdout), "from stdin");
    }

    #[test]
    fn a_timeout_kills_the_process_promptly() {
        let mut s = RunSpec::shell(if cfg!(windows) {
            "ping -n 30 127.0.0.1 > nul"
        } else {
            "sleep 30"
        });
        s.timeout = Duration::from_millis(300);
        let start = Instant::now();
        let r = run(&s).unwrap();
        assert!(r.timed_out && r.code.is_none() && !r.ok());
        assert!(
            start.elapsed() < Duration::from_secs(10),
            "took {:?}",
            start.elapsed()
        );
    }

    #[test]
    fn missing_programs_are_an_error_not_a_panic() {
        assert!(run(&RunSpec::new("definitely-not-a-real-program-xyz", vec![])).is_err());
    }

    #[test]
    fn programs_are_run_without_a_shell_when_given_arguments() {
        // Arguments are not interpreted: the `;` is just text.
        let line = if cfg!(windows) {
            (
                "cmd",
                vec!["/C".to_owned(), "echo".to_owned(), "a;b".to_owned()],
            )
        } else {
            ("echo", vec!["a;b".to_owned()])
        };
        let r = run(&RunSpec::new(line.0, line.1)).unwrap();
        assert!(text(&r.stdout).contains("a;b"), "{:?}", text(&r.stdout));
    }
}
