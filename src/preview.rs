use std::path::Path;

use anyhow::{Context, Result};
use windows::Win32::Graphics::Direct2D::Common::{
    D2D1_ALPHA_MODE_PREMULTIPLIED, D2D1_COLOR_F, D2D1_PIXEL_FORMAT,
};
use windows::Win32::Graphics::Direct2D::{
    D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT, D2D1_FACTORY_TYPE_SINGLE_THREADED,
    D2D1_FEATURE_LEVEL_DEFAULT, D2D1_RENDER_TARGET_PROPERTIES, D2D1_RENDER_TARGET_TYPE_DEFAULT,
    D2D1_RENDER_TARGET_USAGE_NONE, D2D1CreateFactory, ID2D1Factory, ID2D1RenderTarget,
};
use windows::Win32::Graphics::DirectWrite::{
    DWRITE_FACTORY_TYPE_SHARED, DWRITE_FONT_STRETCH_NORMAL, DWRITE_FONT_STYLE_NORMAL,
    DWRITE_FONT_WEIGHT_NORMAL, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, DWRITE_TEXT_ALIGNMENT_CENTER,
    DWRITE_WORD_WRAPPING_NO_WRAP, DWriteCreateFactory, IDWriteFactory, IDWriteFactory3,
    IDWriteFactory7, IDWriteFontCollection, IDWriteFontCollection1, IDWriteFontFile,
    IDWriteFontSet, IDWriteFontSetBuilder, IDWriteFontSetBuilder1, IDWriteStringList,
    IDWriteTextFormat, IDWriteTextLayout,
};
use windows::Win32::Graphics::Dxgi::Common::DXGI_FORMAT_B8G8R8A8_UNORM;
use windows::Win32::Graphics::Imaging::{
    CLSID_WICImagingFactory2, GUID_WICPixelFormat32bppPBGRA, IWICBitmap, IWICImagingFactory,
    WICBitmapCacheOnLoad,
};
use windows::Win32::System::Com::{
    CLSCTX_INPROC_SERVER, COINIT_MULTITHREADED, CoCreateInstance, CoInitializeEx, CoUninitialize,
};
use windows::core::{Interface, PCWSTR};
use windows_numerics::Vector2;

use crate::catalog::FontItem;

pub(crate) struct PreviewImage {
    pub rgba: Vec<u8>,
    pub width: u32,
    pub height: u32,
}

struct ComApartment(bool);

impl ComApartment {
    unsafe fn initialize() -> Self {
        Self(unsafe { CoInitializeEx(None, COINIT_MULTITHREADED) }.is_ok())
    }
}

impl Drop for ComApartment {
    fn drop(&mut self) {
        if self.0 {
            unsafe { CoUninitialize() };
        }
    }
}

struct FontResources {
    collection: IDWriteFontCollection,
    family: String,
}

pub(crate) fn render(
    font: &FontItem,
    text: &str,
    background: [u8; 3],
    text_color: [u8; 3],
    width: u32,
    height: u32,
    font_size: f32,
) -> Result<PreviewImage> {
    unsafe {
        let _apartment = ComApartment::initialize();
        let d2d: ID2D1Factory = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, None)
            .context("Direct2D factory creation failed")?;
        let dwrite: IDWriteFactory7 = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED)
            .context("DirectWrite factory creation failed")?;
        let wic: IWICImagingFactory =
            CoCreateInstance(&CLSID_WICImagingFactory2, None, CLSCTX_INPROC_SERVER)
                .context("WIC factory creation failed")?;
        let resources = if let Some(path) = font.path.as_deref() {
            load_local_font(&dwrite, path)?
        } else {
            load_system_font(&dwrite, &font.family_name)?
        };
        let layout = create_layout(&dwrite, &resources, text, width, height, font_size)?;
        draw(&d2d, &wic, &layout, background, text_color, width, height)
    }
}

unsafe fn load_system_font(factory: &IDWriteFactory7, family: &str) -> Result<FontResources> {
    let base: IDWriteFactory = factory.cast()?;
    let mut collection = None;
    unsafe { base.GetSystemFontCollection(&mut collection, false)? };
    Ok(FontResources {
        collection: collection.context("system font collection was null")?,
        family: family.to_string(),
    })
}

unsafe fn load_local_font(factory: &IDWriteFactory7, path: &Path) -> Result<FontResources> {
    let path = wide(&path.to_string_lossy());
    let file: IDWriteFontFile =
        unsafe { factory.CreateFontFileReference(PCWSTR(path.as_ptr()), None)? };
    let factory3: IDWriteFactory3 = factory.cast()?;
    let builder: IDWriteFontSetBuilder = unsafe { factory3.CreateFontSetBuilder()? };
    let builder: IDWriteFontSetBuilder1 = builder.cast()?;
    unsafe { builder.AddFontFile(&file)? };
    let set: IDWriteFontSet = unsafe { builder.CreateFontSet()? };
    let collection: IDWriteFontCollection1 =
        unsafe { factory3.CreateFontCollectionFromFontSet(&set)? };
    Ok(FontResources {
        collection: collection.cast()?,
        family: unsafe { font_set_family_name(&set) }.unwrap_or_else(|| "Segoe UI".to_string()),
    })
}

unsafe fn font_set_family_name(set: &IDWriteFontSet) -> Option<String> {
    use windows::Win32::Graphics::DirectWrite::DWRITE_FONT_PROPERTY_ID_FAMILY_NAME;
    let strings: IDWriteStringList = unsafe {
        set.GetPropertyValues(DWRITE_FONT_PROPERTY_ID_FAMILY_NAME)
            .ok()?
    };
    let length = unsafe { strings.GetStringLength(0).ok()? } as usize;
    let mut buffer = vec![0u16; length + 1];
    unsafe { strings.GetString(0, &mut buffer).ok()? };
    Some(String::from_utf16_lossy(&buffer[..length]))
}

unsafe fn create_layout(
    factory: &IDWriteFactory7,
    font: &FontResources,
    text: &str,
    width: u32,
    height: u32,
    font_size: f32,
) -> Result<IDWriteTextLayout> {
    let family = wide(&font.family);
    let locale = wide("ja-JP");
    let base: IDWriteFactory = factory.cast()?;
    let format: IDWriteTextFormat = unsafe {
        base.CreateTextFormat(
            PCWSTR(family.as_ptr()),
            &font.collection,
            DWRITE_FONT_WEIGHT_NORMAL,
            DWRITE_FONT_STYLE_NORMAL,
            DWRITE_FONT_STRETCH_NORMAL,
            font_size,
            PCWSTR(locale.as_ptr()),
        )?
    };
    unsafe {
        format.SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER)?;
        format.SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER)?;
        format.SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP)?;
    }
    let text = text.encode_utf16().collect::<Vec<_>>();
    Ok(unsafe { base.CreateTextLayout(&text, &format, width as f32, height as f32)? })
}

unsafe fn draw(
    d2d: &ID2D1Factory,
    wic: &IWICImagingFactory,
    layout: &IDWriteTextLayout,
    background: [u8; 3],
    text_color: [u8; 3],
    width: u32,
    height: u32,
) -> Result<PreviewImage> {
    let bitmap: IWICBitmap = unsafe {
        wic.CreateBitmap(
            width,
            height,
            &GUID_WICPixelFormat32bppPBGRA,
            WICBitmapCacheOnLoad,
        )?
    };
    let properties = D2D1_RENDER_TARGET_PROPERTIES {
        r#type: D2D1_RENDER_TARGET_TYPE_DEFAULT,
        pixelFormat: D2D1_PIXEL_FORMAT {
            format: DXGI_FORMAT_B8G8R8A8_UNORM,
            alphaMode: D2D1_ALPHA_MODE_PREMULTIPLIED,
        },
        dpiX: 96.0,
        dpiY: 96.0,
        usage: D2D1_RENDER_TARGET_USAGE_NONE,
        minLevel: D2D1_FEATURE_LEVEL_DEFAULT,
    };
    let target: ID2D1RenderTarget =
        unsafe { d2d.CreateWicBitmapRenderTarget(&bitmap, &properties)? };
    let brush = unsafe {
        target.CreateSolidColorBrush(
            &D2D1_COLOR_F {
                r: f32::from(text_color[0]) / 255.0,
                g: f32::from(text_color[1]) / 255.0,
                b: f32::from(text_color[2]) / 255.0,
                a: 1.0,
            },
            None,
        )?
    };
    let clear = D2D1_COLOR_F {
        r: f32::from(background[0]) / 255.0,
        g: f32::from(background[1]) / 255.0,
        b: f32::from(background[2]) / 255.0,
        a: 1.0,
    };
    unsafe {
        target.BeginDraw();
        target.Clear(Some(&clear));
        target.DrawTextLayout(
            Vector2::new(0.0, 0.0),
            layout,
            &brush,
            D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT,
        );
        target.EndDraw(None, None)?;
    }

    let stride = width * 4;
    let mut rgba = vec![0; stride as usize * height as usize];
    unsafe { bitmap.CopyPixels(std::ptr::null(), stride, &mut rgba)? };
    for pixel in rgba.chunks_exact_mut(4) {
        pixel.swap(0, 2);
    }
    Ok(PreviewImage {
        rgba,
        width,
        height,
    })
}

fn wide(value: &str) -> Vec<u16> {
    value.encode_utf16().chain(std::iter::once(0)).collect()
}
