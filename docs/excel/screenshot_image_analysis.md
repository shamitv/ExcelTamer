# Screenshot Capture & Image Analysis

ExcelTamer can capture visual screenshots of Excel sheets or ranges and analyse them using a vision-capable LLM. This is useful for understanding formatting, charts, conditional formatting, and other visual elements that are not captured by value-only reads.

---

## Screenshot Capture

**Core method:** `ExcelAutomation.capture_screenshot_png(sheet_name, output_path, cell_range=None)`

**File:** [`ExcelAutomation.py`](file:///d:/work/ExcelTamer/ExcelTamer/ExcelAutomation.py)

```python
success = automation.capture_screenshot_png(
    sheet_name="Dashboard",
    output_path="C:/tmp/dashboard.png",
    cell_range="A1:H20"     # optional — defaults to used range
)
```

### How It Works

1. **Select range**: If `cell_range` is not provided, it defaults to `sheet.used_range.address`.
2. **Scroll into view**: Calls `sheet.range(cell_range).api.Show()` to ensure the range is visible in Excel.
3. **Export to PNG**: Uses xlwings' `range.to_png(output_path)` which leverages Excel's COM APIs to render the range as an image.
4. **Returns** `True` on success, `False` on failure (with an error printed).

> [!NOTE]
> `to_png()` captures the range exactly as it appears in Excel — including cell colours, fonts, borders, charts overlapping the range, and conditional formatting.

---

## Image Analysis via LLM

**LangChain tool:** `ExcelAnalyzeImageTool` (`excel_analyze_image`)

**File:** [`ExcelTamerTools.py`](file:///d:/work/ExcelTamer/ExcelTamer/ExcelTamerAgent/ExcelTamerTools.py)

This tool combines screenshot capture with a vision LLM to answer questions about the visual content of a spreadsheet.

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `question` | `str` | Yes | The question to answer about the image |
| `sheet_name` | `str` | Yes | Sheet to capture |
| `cell_range` | `str` | No | Specific range to capture (whole sheet if omitted) |

### Workflow

```
User question + sheet/range
        │
        ▼
┌─────────────────────┐
│  take_screenshot()   │    ← runs on executor thread (COM)
│  - capture to temp   │
│  - read as base64    │
│  - delete temp file  │
└────────┬────────────┘
         │  data:image/png;base64,...
         ▼
┌─────────────────────────────┐
│  ask_question_about_image() │
│  - build HumanMessage with  │
│    image_url + question     │
│  - invoke vision LLM       │
└────────┬────────────────────┘
         │
         ▼
    LLM response (text)
```

### Implementation Details

1. **Temporary file**: A `tempfile.NamedTemporaryFile(suffix=".png")` is created, the screenshot is saved there, read back as base64, and the file is immediately deleted.

2. **LLM message format**: The image is sent as a `data:image/png;base64,...` URL in a `HumanMessage`:

   ```python
   HumanMessage(content=[
       {"type": "text", "text": "Please provide a concise response..."},
       {"type": "image_url", "image_url": {"url": data_url}}
   ])
   ```

3. **Threading**: Screenshot capture runs on the shared `ThreadPoolExecutor` (see [Threading Model](threading_model.md)) to ensure COM calls happen on the correct thread. The LLM invocation runs on the calling thread (it's a network call, not a COM call).

---

## Availability

| Interface | Available? | Tool Name |
|---|---|---|
| LangChain Agent | ✅ | `excel_analyze_image` |
| MCP Server | ❌ | Not yet exposed as an MCP tool |

> [!TIP]
> For MCP clients, use `excel.read_range` or `excel.read_sheet_preview` for data inspection. Image analysis requires the LangChain Agent interface.
