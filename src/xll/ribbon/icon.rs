//! Status icons for the ribbon button: a flat coloured disc, rendered into a 32-bit DIB and
//! wrapped as an `IPictureDisp` (what the ribbon's `getImage` callback expects).

use std::ffi::c_void;

use windows::Win32::Graphics::Gdi::{
    BI_RGB, BITMAPINFO, BITMAPINFOHEADER, CreateDIBSection, DIB_RGB_COLORS, HBITMAP, HDC,
};
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Ole::{
    OleCreatePictureIndirect, PICTDESC, PICTDESC_0, PICTDESC_0_0, PICTYPE_BITMAP,
};
use windows::core::Result;

use crate::ui::Light;

const SIZE: i32 = 32;

fn rgb(light: Light) -> (u8, u8, u8) {
    match light {
        Light::Green => (46, 160, 67),
        Light::Yellow => (245, 190, 0),
        Light::Red => (214, 48, 49),
    }
}

/// BGRA pixels (straight alpha), top-down: an anti-aliased disc with a darker rim.
pub fn render(light: Light) -> Vec<u8> {
    let (r, g, b) = rgb(light);
    let c = (SIZE as f32 - 1.0) / 2.0;
    let outer = SIZE as f32 / 2.0 - 1.5;
    let rim = outer - 2.0;
    let mut px = vec![0u8; (SIZE * SIZE * 4) as usize];
    for y in 0..SIZE {
        for x in 0..SIZE {
            let d = ((x as f32 - c).powi(2) + (y as f32 - c).powi(2)).sqrt();
            // 1px-wide linear falloff at the outer edge.
            let alpha = (outer + 0.5 - d).clamp(0.0, 1.0);
            let shade = if d > rim { 0.72 } else { 1.0 };
            let i = ((y * SIZE + x) * 4) as usize;
            px[i] = (b as f32 * shade) as u8;
            px[i + 1] = (g as f32 * shade) as u8;
            px[i + 2] = (r as f32 * shade) as u8;
            px[i + 3] = (alpha * 255.0) as u8;
        }
    }
    px
}

/// Creates a fresh picture each call; the picture owns (and frees) its bitmap.
pub fn picture(light: Light) -> Result<IDispatch> {
    let pixels = render(light);
    let bmi = BITMAPINFO {
        bmiHeader: BITMAPINFOHEADER {
            biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
            biWidth: SIZE,
            biHeight: -SIZE, // negative = top-down
            biPlanes: 1,
            biBitCount: 32,
            biCompression: BI_RGB.0,
            ..Default::default()
        },
        ..Default::default()
    };
    let mut bits: *mut c_void = std::ptr::null_mut();
    // SAFETY: standard DIB-section creation; `bits` receives a buffer of SIZE*SIZE*4 bytes that
    // lives as long as the bitmap.
    let hbitmap: HBITMAP = unsafe {
        CreateDIBSection(
            Some(HDC::default()),
            &bmi,
            DIB_RGB_COLORS,
            &mut bits,
            None,
            0,
        )?
    };
    // SAFETY: see above; lengths match.
    unsafe { std::ptr::copy_nonoverlapping(pixels.as_ptr(), bits.cast::<u8>(), pixels.len()) };

    let desc = PICTDESC {
        cbSizeofstruct: std::mem::size_of::<PICTDESC>() as u32,
        picType: PICTYPE_BITMAP.0 as u32,
        Anonymous: PICTDESC_0 {
            bmp: PICTDESC_0_0 {
                hbitmap,
                hpal: Default::default(),
            },
        },
    };
    // SAFETY: `desc` is valid; fOwn = TRUE hands bitmap ownership to the picture.
    unsafe { OleCreatePictureIndirect::<IDispatch>(&desc, true) }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn disc_is_opaque_in_the_middle_and_transparent_in_the_corner() {
        let px = render(Light::Red);
        let at = |x: i32, y: i32| &px[((y * SIZE + x) * 4) as usize..][..4];
        assert_eq!(at(16, 16), &[49, 48, 214, 255]); // BGRA of (214, 48, 49)
        assert_eq!(at(0, 0)[3], 0);
        assert_ne!(render(Light::Green), render(Light::Yellow));
    }
}
