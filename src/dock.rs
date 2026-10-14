//! Geometry for docking the Lua files panel into Excel's window.
//!
//! Excel lays out its workbook area (the `XLDESK` child of the main frame) by itself and
//! re-applies full width whenever the frame is resized, the ribbon collapses, the formula bar
//! grows and so on. We carve out a strip at its right edge: shrink `XLDESK` and park the panel
//! in the freed strip. Because Excel can undo that at any time, [`Dock::plan`] is run on a short
//! timer and works out, from what it sees, whether Excel re-laid out (then the full width is
//! whatever `XLDESK` now has) or whether `XLDESK` is still as we left it.
//!
//! All coordinates are in the frame's client space. Pure, so it is tested without Excel.

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Rect {
    pub left: i32,
    pub top: i32,
    pub right: i32,
    pub bottom: i32,
}

/// What to apply: the new workbook-area rectangle and the panel's rectangle.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Layout {
    pub desk: Rect,
    pub panel: Rect,
}

pub const MIN_PANEL_WIDTH: i32 = 220;
/// The panel never takes more than this share of the workbook area's width.
const MAX_SHARE_PERCENT: i32 = 60;

#[derive(Clone, Copy, Debug, Default)]
pub struct Dock {
    /// Right edge of the workbook area as Excel last laid it out (before we shrank it).
    natural_right: Option<i32>,
    /// Right edge we last applied; if we still see it, Excel has not re-laid out since.
    applied_right: Option<i32>,
    frame_width: Option<i32>,
    /// Consecutive plans in which Excel had put `XLDESK` back to the same full width.
    resets: u32,
}

/// Excel undoing our change on this many consecutive ticks means it is actively fighting us.
const FIGHTING_AFTER: u32 = 40;

impl Dock {
    pub fn new() -> Self {
        Dock::default()
    }

    /// Clamps a requested panel width to something sensible for a workbook area `avail` wide.
    pub fn clamp_width(want: i32, avail: i32) -> i32 {
        let max = (avail * MAX_SHARE_PERCENT / 100).max(MIN_PANEL_WIDTH.min(avail / 2));
        want.clamp(MIN_PANEL_WIDTH.min(avail / 2), max)
    }

    /// `desk` is where `XLDESK` is right now; `frame_width` the frame's client width.
    pub fn plan(&mut self, desk: Rect, frame_width: i32, want_width: i32) -> Layout {
        let same_frame = self.frame_width == Some(frame_width);
        let relaid = self.applied_right != Some(desk.right) || !same_frame;
        // Excel putting the very same full width back, tick after tick, is a fight.
        if relaid && same_frame && self.natural_right == Some(desk.right) {
            self.resets += 1;
        } else {
            self.resets = 0;
        }
        if relaid || self.natural_right.is_none() {
            self.natural_right = Some(desk.right);
        }
        self.frame_width = Some(frame_width);
        let natural = self.natural_right.unwrap_or(desk.right);
        let width = Self::clamp_width(want_width, natural - desk.left);
        let new_right = natural - width;
        self.applied_right = Some(new_right);
        Layout {
            desk: Rect {
                right: new_right,
                ..desk
            },
            panel: Rect {
                left: new_right,
                top: desk.top,
                right: natural,
                bottom: desk.bottom,
            },
        }
    }

    /// True when Excel has undone our layout on every recent tick (a flicker loop): docking is not
    /// workable on this Excel and the panel should float instead.
    pub fn is_fighting(&self) -> bool {
        self.resets >= FIGHTING_AFTER
    }

    /// Where `XLDESK`'s right edge should go when the panel goes away, if it is still as we left
    /// it (if Excel re-laid out in the meantime it is already right).
    pub fn restore_right(&self, current_right: i32) -> Option<i32> {
        match (self.applied_right, self.natural_right) {
            (Some(applied), Some(natural)) if applied == current_right && natural != applied => {
                Some(natural)
            }
            _ => None,
        }
    }

    pub fn reset(&mut self) {
        *self = Dock::default();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn desk(right: i32) -> Rect {
        Rect {
            left: 0,
            top: 120,
            right,
            bottom: 800,
        }
    }

    #[test]
    fn carves_a_strip_off_the_right_edge() {
        let mut d = Dock::new();
        let l = d.plan(desk(1600), 1600, 340);
        assert_eq!(l.desk, desk(1260));
        assert_eq!(
            l.panel,
            Rect {
                left: 1260,
                top: 120,
                right: 1600,
                bottom: 800
            }
        );
    }

    #[test]
    fn leaves_things_alone_when_excel_has_not_moved_anything() {
        let mut d = Dock::new();
        let first = d.plan(desk(1600), 1600, 340);
        // Next tick: XLDESK is exactly where we put it, so the plan is unchanged.
        assert_eq!(d.plan(first.desk, 1600, 340), first);
    }

    #[test]
    fn follows_excel_when_it_relays_out_at_full_width() {
        let mut d = Dock::new();
        d.plan(desk(1600), 1600, 340);
        // Excel re-applied full width (e.g. the formula bar grew): shrink again.
        let l = d.plan(
            Rect {
                top: 150,
                ..desk(1600)
            },
            1600,
            340,
        );
        assert_eq!((l.desk.right, l.desk.top, l.panel.left), (1260, 150, 1260));
    }

    #[test]
    fn follows_the_frame_when_it_is_resized() {
        let mut d = Dock::new();
        d.plan(desk(1600), 1600, 340);
        let wider = d.plan(desk(1900), 1900, 340);
        assert_eq!((wider.desk.right, wider.panel.right), (1560, 1900));
        // A shrink by more than the panel width must not be mistaken for "still as we left it".
        let narrower = d.plan(desk(1000), 1000, 340);
        assert_eq!(
            (
                narrower.desk.right,
                narrower.panel.left,
                narrower.panel.right
            ),
            (660, 660, 1000)
        );
    }

    #[test]
    fn a_frame_resize_that_lands_exactly_on_our_edge_is_still_noticed() {
        let mut d = Dock::new();
        d.plan(desk(1600), 1600, 340); // applied_right = 1260
        // Frame shrinks so Excel's full width happens to equal what we applied last time.
        let l = d.plan(desk(1260), 1260, 340);
        assert_eq!(l.panel.right, 1260);
        assert_eq!(l.desk.right, 920);
    }

    #[test]
    fn width_is_clamped_for_small_and_huge_requests() {
        assert_eq!(Dock::clamp_width(50, 1600), MIN_PANEL_WIDTH);
        assert_eq!(Dock::clamp_width(5000, 1000), 600);
        // A tiny workbook area still leaves room for both sides.
        assert_eq!(Dock::clamp_width(340, 300), 180);
        assert_eq!(
            Dock::clamp_width(340, 100),
            60,
            "never more than 60% of a very small area"
        );
    }

    #[test]
    fn a_one_off_relayout_is_normal_but_a_relayout_every_tick_is_a_fight() {
        let mut d = Dock::new();
        let first = d.plan(desk(1600), 1600, 340);
        d.plan(desk(1600), 1600, 340); // Excel reset once...
        d.plan(first.desk, 1600, 340); // ...then left it alone
        assert!(!d.is_fighting());
        for _ in 0..FIGHTING_AFTER + 1 {
            d.plan(desk(1600), 1600, 340); // Excel resets before every single tick
        }
        assert!(d.is_fighting());
        // A real frame resize is not a fight, and calm ticks clear the count.
        d.plan(desk(1900), 1900, 340);
        assert!(!d.is_fighting());
    }

    #[test]
    fn restore_gives_the_width_back_only_if_excel_has_not_already() {
        let mut d = Dock::new();
        let l = d.plan(desk(1600), 1600, 340);
        assert_eq!(d.restore_right(l.desk.right), Some(1600));
        assert_eq!(
            d.restore_right(1600),
            None,
            "Excel already re-laid out at full width"
        );
        let fresh = Dock::new();
        assert_eq!(fresh.restore_right(1234), None);
    }
}
