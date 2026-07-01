# NI USB-6002 + INA126 + Load Cell + Servo Monitor

Real-time Python desktop app for reading a load cell via an INA126 instrumentation amplifier
and an NI USB-6002 DAQ, with Arduino-driven servo control and waveform sweep.

---

## Features

- **2.5 V reference** held on AO0 for the full session
- **Dual readout** — raw voltage (V) and calibrated weight side by side
- **Two-point calibration** (tare + known weight) with auto-save to `calibration.json`
- **Live tare offset** — zero the weight display at any time without changing calibration
- **Selectable graph Y-axis** — Voltage, kg, g, lb, oz, or N
- **Running Min / Max / Avg** (in the active graph unit)
- **Servo control via Arduino UNO** — manual slider, preset angle buttons (0°/45°/90°/135°/180°)
- **Waveform sweep** — drive the servo with Sine, Square, Triangle, Sawtooth, or Rev. Sawtooth at adjustable frequency, amplitude, and centre angle
- **Servo angle graph** — second scrolling plot showing angle vs time
- **Excel recording** — capture time, voltage, weight, and servo angle to a `.xlsx` file
- **Filter tool** (`filter_tool.py`) — post-process recordings with a digital Butterworth filter, zoom/pan plots, analyse dynamic vs quasi-static regions, and export publication-ready images
- **Demo mode** — works without NI hardware using a simulated signal

---

## Requirements

### Software
- Python 3.9 or later
- NI-DAQmx driver — [ni.com/downloads](https://www.ni.com/en/support/downloads/drivers/download.ni-daq-mx.html)
- Arduino IDE — to upload the servo sketch

### Python packages
```
pip install -r requirements.txt
```

---

## Hardware connections

### Load cell / INA126
```
NI USB-6002        INA126
─────────────      ──────────────────────────────────
AO0               → REF (pin 5)     2.5 V reference
AO GND            → GND
AI0+              ← OUT (pin 6)     Amplified load cell signal
AI GND              (internal reference)
```

### Servo / Arduino UNO
```
Arduino UNO        Servo
───────────        ────────────────────────────────────
Pin 9            → Signal (orange/yellow)
GND              → GND    (brown/black)   shared with power supply GND
                   Power  (red)         ← External 5 V supply
                                          Do NOT use the Arduino 5V pin under servo load
```
Connect the Arduino to the PC via USB. The app communicates with it over serial (default **COM3**).

---

## Setup

### 1 — Upload the Arduino sketch
Open `servo_controller/servo_controller.ino` in the Arduino IDE and upload it to the UNO.
The servo will move to 90° on power-up and print `READY` on the serial monitor.

### 2 — Install Python dependencies
```bash
pip install -r requirements.txt
```

### 3 — Run the app
```bash
python main.py
```

---

## Live tare

The weight readout has **Tare** and **Clear tare** buttons below the value display.

- **Tare** — averages the last 10 samples and stores the current weight as the zero offset. A yellow label shows the active offset value.
- **Clear tare** — resets the offset to zero.

The offset is applied to the weight readout, the scrolling graph, and any Excel recording in progress. It does not change the two-point calibration stored in `calibration.json`.

---

## Servo control

### Connecting
1. In the **SERVO CONTROL** panel, click the **Arduino port** dropdown — it auto-lists available COM ports.
2. Select the port matching your Arduino (check Device Manager if unsure).
3. Click **Connect**. The status label turns green when connected.

### Manual control
- Drag the **Angle** slider or click a preset button (**0° / 45° / 90° / 135° / 180°**).

### Waveform sweep
Configure in the **WAVEFORM** section of the servo panel:

| Control | Description |
|---------|-------------|
| Shape | Sine, Square, Triangle, Sawtooth, Rev. Sawtooth |
| Freq | Sweep frequency in Hz (0.05–5.0) |
| Center | Midpoint angle the waveform oscillates around (0–180°) |
| Amplitude | How far from centre the sweep extends (0–90°) |
| ▶ Start / ■ Stop | Starts or stops the sweep |

While the waveform is running, the manual slider and preset buttons are locked. The servo angle graph updates live.

---

## Calibration procedure

The app uses two-point linear calibration. Open the dialog with the **Calibrate…** button.

### Step 1 — Tare (zero)
1. Remove all weight from the load cell.
2. Wait for the voltage reading to stabilise.
3. Click **Tare / Zero**. The app averages 1 second of samples and saves the zero voltage immediately.

### Step 2 — Set span
1. Place a **known reference weight** on the load cell.
2. Enter the exact weight and unit in the dialog.
3. Wait for the reading to stabilise.
4. Click **Set Span**. The span voltage and weight are saved immediately.

Calibration is written to `calibration.json` next to `main.py` and loaded automatically on the next launch.

### Calibration formula
```
weight = (voltage − zero_V) / (span_V − zero_V) × known_weight
```

---

## Recording

1. Click **● Record** in the status bar and choose a save path.
2. Data is captured continuously until you click **■ Stop**.
3. The recording is saved as an `.xlsx` file with columns: **Time (s)**, **Voltage (V)**, **Weight**, **Servo Angle (°)**.

If the window is closed while recording is active the file is saved automatically.

---

## Filter tool

`filter_tool.py` is a standalone post-processing tool for recordings made with `main.py`.

```bash
python filter_tool.py                   # opens file dialog
python filter_tool.py recording.xlsx    # loads file directly
```

### Filter settings

| Control | Description |
|---------|-------------|
| Column | Select which data column to filter (Voltage, Weight, etc.) |
| Type | Low-pass, High-pass, Band-pass, or Band-stop |
| Order | Filter order 1–8 |
| Cutoff | Cutoff frequency in Hz; band types show two sliders |
| Offset | Constant added to the filtered output; unit label shows the active column |
| Auto | Re-applies the filter automatically as parameters change |
| Remove DC | Subtracts the signal mean before filtering to avoid transient artifacts |
| Invert | Negates the filtered signal (useful for reversed-polarity sensors) |
| Apply | Runs the filter manually when Auto is off |
| Save filtered | Saves a new `.xlsx` with a `[filtered]` column added |

Both the signal plot and the Bode magnitude plot have a **navigation toolbar** for zoom, pan, and view history.

### Signal analysis

The **SIGNAL ANALYSIS** panel identifies dynamic and quasi-static regions in the filtered signal.

| Control | Description |
|---------|-------------|
| Window | Smoothing window for the activity measure (seconds) |
| Threshold | Fraction of peak derivative used to separate dynamic from quasi-static (0–1) |
| Min value | Exclude regions whose peak falls below this value (removes noise-floor detections) |
| Max regions | Keep only the top N regions by peak value; blank keeps all |
| Analyse | Runs the classification and shades regions on the plot |
| Export image… | Saves a PNG/PDF/SVG of the signal plot with all annotations |
| Clear | Removes shading and clears the results table |
| ? | Opens a help window explaining every metric and analysis control |

After clicking **Analyse**:
- **Yellow shading** marks dynamic regions (signal changing rapidly)
- **Green shading** marks quasi-static regions (signal settling or holding steady)
- A results table below shows per-region statistics: t-start, t-end, duration, max, min, mean, std dev, and peak-to-peak
- A summary line shows the total dynamic and quasi-static time

### Exporting an image

**Export image…** produces a standalone figure file that includes:
- The signal plot with Original and Filtered lines, region shading, and per-region labels (type + mean)
- Title showing filename, column, filter type, order, cutoff frequency, sample rate, and offset (if non-zero)
- A styled analysis table (if Analyse has been run) with colour-coded rows and a summary line

Supported formats: **PNG** (150 dpi), **PDF**, **SVG**.

---

## Configuration

Edit the constants at the top of `main.py`:

| Constant | Default | Description |
|----------|---------|-------------|
| `DEVICE_NAME` | `"Dev1"` | NI MAX device name |
| `REFERENCE_VOLTAGE` | `2.5` | AO0 output voltage |
| `SAMPLE_RATE` | `100` | AI sample rate (Sa/s) |
| `CAL_AVG_SAMPLES` | `100` | Samples averaged at tare/span (= 1 s) |
| `HISTORY_SECONDS` | `15` | Graph rolling window |
| `SERIAL_PORT` | `"COM3"` | Arduino COM port |
| `SERIAL_BAUD` | `9600` | Serial baud rate (must match sketch) |

To find your NI device name open **NI MAX → Devices and Interfaces**.  
To find your Arduino port open **Device Manager → Ports (COM & LPT)**.

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| App opens in DEMO mode | NI-DAQmx driver not installed, or device unplugged |
| `DaqError: device not found` | Check NI MAX; match `DEVICE_NAME` to what is shown there |
| `ModuleNotFoundError: nidaqmx` | `pip install nidaqmx` |
| `ModuleNotFoundError: serial` | `pip install pyserial` |
| `ModuleNotFoundError: openpyxl` | `pip install openpyxl` |
| `ModuleNotFoundError: scipy` | `pip install scipy` — ensure you use the same Python that runs the script |
| Servo not responding | Check COM port in Device Manager; click Connect again |
| Servo jitters or makes noise | Ensure servo power comes from an external supply, not Arduino 5V |
| Weight reads wrong after loading | Re-run calibration with a fresh tare |
| Reading drifts over time | Temperature effect on load cell — re-tare before each session |

---

## License

MIT
