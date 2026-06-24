/*
  servo_controller.ino
  Receives angle commands from Python over USB serial and moves a servo.

  Wiring:
    Servo signal (orange/yellow) -> Pin 9
    Servo power  (red)           -> External 5 V supply (do NOT use Arduino 5V pin
                                    under load — it can brown-out the board)
    Servo GND    (brown/black)   -> GND (shared with Arduino and power supply)

  Protocol (9600 baud, newline-terminated):
    Python sends:  "90\n"
    Arduino moves servo to 90° and replies: "OK:90\n"

  Valid range: 0–180 degrees. Out-of-range values are clamped automatically.
*/

#include <Servo.h>

const int SERVO_PIN    = 9;
const int BAUD_RATE    = 9600;
const int ANGLE_MIN    = 0;
const int ANGLE_MAX    = 180;
const int ANGLE_START  = 90;

Servo servo;
String inputBuffer = "";

void setup() {
  Serial.begin(BAUD_RATE);
  servo.attach(SERVO_PIN);
  servo.write(ANGLE_START);
  Serial.println("READY");
}

void loop() {
  while (Serial.available() > 0) {
    char c = (char)Serial.read();

    if (c == '\n') {
      inputBuffer.trim();
      if (inputBuffer.length() > 0) {
        int angle = inputBuffer.toInt();
        angle = constrain(angle, ANGLE_MIN, ANGLE_MAX);
        servo.write(angle);
        Serial.print("OK:");
        Serial.println(angle);
      }
      inputBuffer = "";
    } else if (c != '\r') {
      inputBuffer += c;
    }
  }
}
