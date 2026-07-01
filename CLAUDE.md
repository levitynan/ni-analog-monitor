# NI USB-6002 / INA126 / Load Cell / Servo Monitor — Claude Context

## What this project does
Real-time Python GUI that:
- Outputs a fixed **2.5 V reference** on **AO0** (held open for the lifetime of the process)
- Continuously reads **AI0** (INA126 output) at 100 Sa/s — displays voltage, calibrated weight, scrolling waveform
- Supports **two-point load cell calibration** (tare + span) with JSON persistence and unit selection
- Supports a **live tare offset** on the weight readout (independent of calibration)
- Controls a **servo motor** via an **Arduino UNO** over USB serial — manual slider, preset buttons, and automatic waveform sweep
- Records time/voltage/weight/servo-angle to an **Excel file** (.xlsx) on demand
- Provides a separate **filter tool** (`filter_tool.py`) for post-processing recordings with a digital Butterworth filter

## Key files
| File | Purpose |
|------|---------|
| `main.py` | Single-file app — all hardware, UI, calibration, recording, and waveform logic |
| `filter_tool.py` | Standalone post-processing tool — load xlsx recordings, apply Butterworth filter, analyse regions, export image |
| `servo_controller/servo_controller.ino` | Arduino sketch — receives angle integers over serial, drives servo on pin 9 |
| `calibration.json` | Saved calibration (auto-created on first save) |
| `requirements.txt` | Python dependencies (includes scipy for filter_tool.py) |

## Hardware
- **NI USB-6002 DAQ**
  - **AO0** — 2.5 V DC reference for INA126 REF pin (0–5 V, 12-bit)
  - **AI0** — INA126 output, RSE mode, ±10 V range, 100 Sa/s
  - Device name defaults to `Dev1`
- **Arduino UNO** — connected via USB, generates hardware servo PWM on **pin 9**
  - Receives newline-terminated integer angles over serial (`"90\n"`)
  - Replies `"OK:90\n"` (visible in Arduino Serial Monitor)
  - Serial port defaults to `COM3` — change `SERIAL_PORT` at top of `main.py`
- **Servo motor** — signal wire to Arduino pin 9, powered from an external 5 V supply

## Architecture

### Threading model
| Thread | What it does |
|--------|-------------|
| **UI thread** | Tkinter main loop; `App._tick` fires every 100 ms via `root.after` |
| **Hardware thread** (`_hardware_thread`) | Opens AO task (2.5 V), runs AI continuous acquisition; pushes samples into `_voltage_buf`/`_time_buf` deques behind `self._lock`. Falls back to `_demo_loop` (sine + noise) if NI-DAQmx is absent or device not found. |
| **Waveform thread** (`_waveform_thread_fn`) | Runs at 25 Hz when waveform is active; computes servo angle from selected waveform shape and sends it over serial. Uses `time.perf_counter`-anchored timing to prevent drift. |

Serial writes from the waveform thread and from manual UI controls both call `_send_angle()`, which writes `"{angle}\n"` to `self._serial_port`. The serial port is opened/closed manually via the Connect button (not automatically on startup).

### Calibration (`Calibration` dataclass)
- Two-point linear: `weight = (voltage - zero_v) / (span_v - zero_v) * known_weight`
- `zero_v` — voltage at no-load (tare step, averaged over `CAL_AVG_SAMPLES = 100` readings)
- `span_v` — voltage at known reference weight (span step, same averaging)
- Auto-saved to `calibration.json` immediately on each Tare and Set Span action
- `Calibration.to_weight_in_unit(voltage, display_unit)` converts through kg as a base: `KG_TO_UNIT` dict

### Live tare offset
`self._weight_tare_kg` (float, default 0.0) is a runtime offset stored in kg.
- **Tare** button in the weight readout: averages the last 10 voltage samples, converts to the calibration unit, stores as kg. Yellow label shows the active offset.
- **Clear tare** resets offset to 0.
- Applied as `w_tare = self._weight_tare_kg * KG_TO_UNIT[cal.unit]` and subtracted from weight display, graph plot data, and recorded weight column — all converted to the correct unit per context.
- Does not affect `calibration.json` or the two-point calibration.

### Graph units
`_graph_unit_var` (StringVar) is set to one of `GRAPH_UNITS = ["Voltage", "kg", "g", "lb", "oz", "N"]` via radio buttons.
- **Voltage** — raw V, 2.5 V reference dashed line visible
- **Weight units** — buffer converted to the selected unit; requires valid calibration

### Waveform generator
`_wave_type_var` selects from: Sine, Square, Triangle, Sawtooth, Rev. Sawtooth.
All shapes normalised to [−1, +1]; final angle = `center + amplitude × wave(phase)`, clamped 0–180°.

| Shape | Formula |
|-------|---------|
| Sine | `sin(2π × phase)` |
| Square | `+1` if phase < 0.5 else `−1` |
| Triangle | `1 − 2 × |2×phase − 1|` |
| Sawtooth | `2×phase − 1` |
| Rev. Sawtooth | `1 − 2×phase` |

Starting the waveform disables the manual slider and preset buttons; stopping re-enables them via `root.after(0, _on_waveform_stopped)` (safe cross-thread Tkinter call).

### Recording
`_toggle_recording` asks for a file path then sets `_recording = True`. Each sample appended in the hardware/demo thread includes `(elapsed_time, voltage, weight, servo_angle)`. The weight value has the live tare offset already applied. On stop, `_save_recording` writes an `.xlsx` file with styled headers using `openpyxl`.

### Teardown (Tkinter after-loop fix)
`_on_close` calls `root.quit()` **before** `root.destroy()`. This exits the Tcl event loop so no pending `after` callbacks can fire and trigger `invalid command name` background errors.

### Demo mode
If `nidaqmx` is not installed or the device is not found, the app runs a synthetic sine + noise signal so the UI, calibration, and waveform can all be tested offline. Serial/Arduino servo still works independently of DAQ mode.

## filter_tool.py

Standalone Tkinter app for post-processing `.xlsx` recordings produced by `main.py`.

### Usage
```
python filter_tool.py [recording.xlsx]
```
The file argument is optional; a file-open dialog appears if omitted.

### Features
- **Column selector** — choose any numeric column from the recording (Voltage, Weight, etc.)
- **Filter types** — Low-pass, High-pass, Band-pass, Band-stop (Butterworth, zero-phase via `filtfilt`)
- **Filter order** — 1–8 via slider
- **Offset** — constant value added to the filtered output; unit label updates to match the selected column
- **Auto-apply** — re-runs filter automatically as parameters change when checked
- **Remove DC** — subtracts the column mean before filtering to prevent transient artifacts from large offsets; status bar shows the removed mean value
- **Invert** — negates the filtered signal (and the displayed original) for sensors wired with reversed polarity
- **Zoom / Pan** — matplotlib `NavigationToolbar2Tk` on both the signal plot and Bode plot; supports zoom-to-rectangle, pan, home/reset, back/forward history
- **Frequency response plot** — Bode magnitude (dB) with −3 dB reference line, computed via `freqz`
- **Save filtered** — writes a new `.xlsx` with an added `[filtered]` column alongside the original data

### Signal analysis panel
- **Window** slider — smoothing window for the activity signal (seconds)
- **Threshold** slider — fraction of normalised peak derivative above which a region is classified as dynamic (0–1)
- **Min value** entry — regions whose peak value falls below this threshold are excluded from results (noise-floor filtering)
- **Max regions** entry — keep only the top N regions by peak value; blank means no limit; retained regions are re-sorted chronologically for display
- **Analyse** button — classifies each sample as dynamic (rapidly changing) or quasi-static (settling/holding) using the smoothed absolute derivative of the filtered signal; draws coloured axvspan shading (yellow = dynamic, green = quasi-static) on the signal plot
- Results text widget shows a colour-coded table: region #, type, t-start, t-end, duration, max, min, mean, std dev, peak-to-peak; summary line shows total dynamic/quasi-static time
- **Export image…** — saves a standalone PNG/PDF/SVG containing the signal plot (with shading, region labels showing type + mean, and filter settings in the title) and, if analysis has been run, a styled matplotlib table of all region statistics below it
- **Clear** — removes shading and clears the results table
- **?** button — opens a modal scrollable help window with colour-coded sections explaining every metric and control in the analysis panel (Region classification, Time metrics, Amplitude metrics, Analysis controls)
- Changing filter parameters automatically clears region shading (regions are stale after signal changes)

### Implementation notes
- `HAS_SCIPY` / `SCIPY_ERR`: scipy is imported inside a `try/except`; if missing, applying the filter shows a dialog with the actual import error to help diagnose Python environment mismatches
- Sample rate (`self._fs`) is detected from median sample interval of the Time column
- DC removal is applied **before** filtering: `input_sig = raw - np.mean(raw)` — this prevents the filter from producing large transient spikes at the start of the signal
- Offset is applied **after** filtering and inversion to the final `self._filtered` array
- `self._plot_input` stores the displayed original signal (after DC removal / inversion) so the export can redraw it without re-running the filter
- `self._last_analysis` stores the last region result list so the export can render the table
- Band-pass/Band-stop types show a second cutoff slider (fc_high); single-cutoff types hide it
- `_schedule_apply()` debounces auto-apply by 250 ms to avoid redundant filter calls while sliders are being dragged
- Status bar is packed with `side=tk.BOTTOM` before the plot frame so it remains visible when plots expand to fill the window

## Running
```
pip install -r requirements.txt
python main.py           # main DAQ + servo app
python filter_tool.py    # post-processing filter tool
```
Upload `servo_controller/servo_controller.ino` to the Arduino UNO via the Arduino IDE before connecting.

## Common configuration changes
| What to change | Constant in `main.py` |
|----------------|----------------------|
| NI device name | `DEVICE_NAME = "Dev1"` |
| AI channel / voltage range | `AI_CHANNEL`, `AI_MIN_V`, `AI_MAX_V` |
| Reference voltage | `REFERENCE_VOLTAGE = 2.5` |
| Sample rate | `SAMPLE_RATE = 100` |
| Calibration averaging | `CAL_AVG_SAMPLES = 100` |
| Graph history length | `HISTORY_SECONDS = 15` |
| Arduino serial port | `SERIAL_PORT = "COM3"` |
| Arduino baud rate | `SERIAL_BAUD = 9600` |

## Dependencies
- `nidaqmx` — NI-DAQmx Python wrapper (requires NI-DAQmx driver from ni.com); optional, falls back to demo mode
- `pyserial` — serial communication with the Arduino UNO
- `matplotlib` — embedded graphs via `TkAgg` backend; `NavigationToolbar2Tk` used in filter_tool.py
- `numpy` — statistics, demo signal, calibration averaging
- `openpyxl` — Excel recording export and filtered result save; optional, warns if absent
- `scipy` — Butterworth filter design, frequency response, and region analysis (`uniform_filter1d`) in `filter_tool.py`; optional in main app
- `tkinter` — standard library GUI
