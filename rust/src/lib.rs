use std::path::PathBuf;

use js_sys::{Array, Object, Reflect, Uint8Array};
use office2pdf::config::{ConvertOptions, Format, PaperSize, PdfStandard, SlideRange};
use office2pdf::{convert_bytes as upstream_convert_bytes, error::ConvertWarning};
use serde::Deserialize;
use serde_wasm_bindgen::from_value;
use wasm_bindgen::prelude::*;

const ERROR_PAGE_RANGE_UNSUPPORTED: &str = "page_range is not supported by upstream office2pdf 0.6";
const ERROR_MEMORY_LIMIT_UNSUPPORTED: &str =
    "memory_limit_mb is not supported by upstream office2pdf 0.6";

#[derive(Default, Deserialize)]
#[serde(default)]
struct JsConvertOptions {
    #[serde(alias = "pageRange")]
    page_range: Option<String>,
    #[serde(alias = "sheetFilter")]
    #[serde(alias = "sheet_filter")]
    sheet_filter: Option<Vec<String>>, // maps to upstream `sheet_names`
    #[serde(alias = "slideRange")]
    slide_range: Option<String>,
    #[serde(alias = "paperSize")]
    paper_size: Option<String>,
    landscape: Option<bool>,
    #[serde(alias = "fontPaths")]
    #[serde(alias = "font_paths")]
    font_paths: Vec<String>,
    #[serde(alias = "pdfStandard")]
    pdf_standard: Option<String>,
    #[serde(alias = "includeWarnings")]
    include_warnings: Option<bool>,
    #[serde(alias = "memoryLimitMb")]
    memory_limit_mb: Option<u64>,
    streaming: Option<bool>,
}

struct ConversionMetrics {
    parse_duration_ms: u64,
    codegen_duration_ms: u64,
    compile_duration_ms: u64,
    total_duration_ms: u64,
    input_size_bytes: u64,
    output_size_bytes: u64,
    page_count: u32,
}

#[wasm_bindgen]
pub fn convert_bytes(data: &[u8], format: &str, options: JsValue) -> Result<JsValue, JsError> {
    let options: JsConvertOptions = if options.is_undefined() || options.is_null() {
        JsConvertOptions::default()
    } else {
        from_value(options).map_err(|error| JsError::new(&error.to_string()))?
    };

    if options.page_range.as_deref().is_some() {
        return Err(JsError::new(ERROR_PAGE_RANGE_UNSUPPORTED));
    }

    if options.memory_limit_mb.is_some() {
        return Err(JsError::new(ERROR_MEMORY_LIMIT_UNSUPPORTED));
    }

    let format = parse_format(format)?;
    let include_warnings = options.include_warnings.unwrap_or(true);

    let options = build_native_options(&options)?;
    let raw = upstream_convert_bytes(data, format, &options)
        .map_err(|error| JsError::new(&error.to_string()))?;

    let warnings = if include_warnings {
        raw.warnings.iter().map(ConvertWarning::to_string).collect()
    } else {
        Vec::new()
    };

    let metrics = raw.metrics.map(|metrics| ConversionMetrics {
        parse_duration_ms: metrics.parse_duration.as_millis() as u64,
        codegen_duration_ms: metrics.codegen_duration.as_millis() as u64,
        compile_duration_ms: metrics.compile_duration.as_millis() as u64,
        total_duration_ms: metrics.total_duration.as_millis() as u64,
        input_size_bytes: metrics.input_size_bytes,
        output_size_bytes: metrics.output_size_bytes,
        page_count: metrics.page_count,
    });

    conversion_result_to_js(raw.pdf, warnings, metrics)
}

fn parse_format(value: &str) -> Result<Format, JsError> {
    match value.trim().to_ascii_lowercase().as_str() {
        "docx" => Ok(Format::Docx),
        "pptx" => Ok(Format::Pptx),
        "xlsx" => Ok(Format::Xlsx),
        _ => Err(JsError::new("format must be one of: docx, pptx, xlsx")),
    }
}

fn parse_sheet_names(raw_names: Option<Vec<String>>) -> Result<Option<Vec<String>>, JsError> {
    raw_names
        .map(|names| {
            if names.iter().any(|name| name.trim().is_empty()) {
                Err(JsError::new("sheet_filter entries must not be empty"))
            } else {
                Ok(names)
            }
        })
        .transpose()
}

fn build_native_options(options: &JsConvertOptions) -> Result<ConvertOptions, JsError> {
    let sheet_names = parse_sheet_names(options.sheet_filter.clone())?;
    let slide_range = match options.slide_range.as_deref() {
        None => None,
        Some(range) => Some(
            SlideRange::parse(range)
                .map_err(|error| JsError::new(format!("invalid slide_range: {error}").as_ref()))?,
        ),
    };

    let paper_size = match options.paper_size.as_deref() {
        None => None,
        Some(size) => Some(
            PaperSize::parse(size)
                .map_err(|error| JsError::new(format!("invalid paper_size: {error}").as_ref()))?,
        ),
    };

    let pdf_standard = match options.pdf_standard.as_deref() {
        None => None,
        Some(value) => Some(parse_pdf_standard(value)?),
    };

    let font_paths = options
        .font_paths
        .iter()
        .map(PathBuf::from)
        .collect::<Vec<_>>();

    Ok(ConvertOptions {
        sheet_names,
        slide_range,
        pdf_standard,
        paper_size,
        font_paths,
        landscape: options.landscape,
        tagged: false,
        pdf_ua: false,
        streaming: options.streaming.unwrap_or(false),
        streaming_chunk_size: None,
    })
}

fn parse_pdf_standard(value: &str) -> Result<PdfStandard, JsError> {
    let normalized = value
        .chars()
        .filter(|character| !matches!(character, '-' | '_' | '/'))
        .flat_map(char::to_lowercase)
        .collect::<String>();

    if normalized == "pdfa2b" {
        Ok(PdfStandard::PdfA2b)
    } else {
        Err(JsError::new("invalid pdf_standard: expected 'pdf/a-2b'"))
    }
}

fn conversion_result_to_js(
    pdf: Vec<u8>,
    warnings: Vec<String>,
    metrics: Option<ConversionMetrics>,
) -> Result<JsValue, JsError> {
    let output = Object::new();
    let pdf_bytes = Uint8Array::from(pdf.as_slice());
    set_property(&output, "pdf", pdf_bytes.as_ref())?;

    let warning_values = Array::new();
    for warning in warnings {
        warning_values.push(&JsValue::from_str(&warning));
    }
    set_property(&output, "warnings", warning_values.as_ref())?;

    match metrics {
        Some(metrics) => {
            let metrics_value = metrics_to_js(metrics)?;
            set_property(&output, "metrics", metrics_value.as_ref())?;
        }
        None => set_property(&output, "metrics", &JsValue::NULL)?,
    }

    Ok(output.into())
}

fn metrics_to_js(metrics: ConversionMetrics) -> Result<Object, JsError> {
    let output = Object::new();
    set_number(&output, "parseDurationMs", metrics.parse_duration_ms)?;
    set_number(&output, "codegenDurationMs", metrics.codegen_duration_ms)?;
    set_number(&output, "compileDurationMs", metrics.compile_duration_ms)?;
    set_number(&output, "totalDurationMs", metrics.total_duration_ms)?;
    set_number(&output, "inputSizeBytes", metrics.input_size_bytes)?;
    set_number(&output, "outputSizeBytes", metrics.output_size_bytes)?;
    set_number(&output, "pageCount", u64::from(metrics.page_count))?;
    Ok(output)
}

fn set_number(target: &Object, key: &str, value: u64) -> Result<(), JsError> {
    set_property(target, key, &JsValue::from_f64(value as f64))
}

fn set_property(target: &Object, key: &str, value: &JsValue) -> Result<(), JsError> {
    let success = Reflect::set(target.as_ref(), &JsValue::from_str(key), value)
        .map_err(|error| JsError::new(&format!("failed to set {key}: {error:?}")))?;

    if success {
        Ok(())
    } else {
        Err(JsError::new(&format!("failed to set {key}")))
    }
}
