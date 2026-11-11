#include <MFRC522.h>
#include <SPI.h>

#define RESET_PIN 9
#define SELECT_PIN 10

MFRC522 rfidReader(SELECT_PIN, RESET_PIN);

int serialCounter = 1;
unsigned long startMillis;
int startHours = 12, startMinutes = 55,
    startSeconds = 50; // Manually set initial time OrderCalendars

struct UserData {
  String cardUID;
  String rollNumber;
  String name;
  String lastTimeIn;
  bool isInside;
};

// Array to store registered users
UserData userList[] = {
    {"2a472f3", "HST - 01", "Vivek", "", false},
    {"29c224b7", "HST - 02", "Ritik", "", false},
    {"d33ff0", "HST - 03", "Ainul", "", false},
};

const int userCount = sizeof(userList) / sizeof(userList[0]);

// Function to get the current time using millis()
String getCurrentTime() {
  unsigned long elapsedSeconds = (millis() - startMillis) / 1000;

  int currentHours = startHours + (elapsedSeconds / 3600);
  int currentMinutes = startMinutes + ((elapsedSeconds % 3600) / 60);
  int currentSeconds = startSeconds + (elapsedSeconds % 60);

  if (currentSeconds >= 60) {
    currentMinutes += currentSeconds / 60;
    currentSeconds %= 60;
  }
  if (currentMinutes >= 60) {
    currentHours += currentMinutes / 60;
    currentMinutes %= 60;
  }
  if (currentHours >= 24) {
    currentHours %= 24;
  } // Reset if over 24 hours

  char formattedTime[20];
  sprintf(formattedTime, "%02d:%02d:%02d", currentHours, currentMinutes,
          currentSeconds);
  return String(formattedTime);
}

// Function to find user index in the array
int findUserIndex(String cardNumber) {
  for (int i = 0; i < userCount; i++) {
    if (userList[i].cardUID.equalsIgnoreCase(cardNumber)) {
      return i;
    }
  }
  return -1; // Not found
}

void setup() {
  Serial.begin(9600);
  while (!Serial && millis() < 2000);
  SPI.begin();
  rfidReader.PCD_Init();
  delay(100);
  startMillis = millis();

  Serial.println(F("========================================"));
  Serial.println(F("   RFID ATTENDANCE SYSTEM READY"));
  Serial.println(F("========================================"));
  Serial.print(F("MFRC522 Status: "));
  rfidReader.PCD_DumpVersionToSerial();
  Serial.println(F("----------------------------------------"));
  Serial.println(F("Serial Number, RFID Card Number, Roll Number, Name, Time In, Time Out, Status"));
}

void loop() {
  if (!rfidReader.PICC_IsNewCardPresent() ||
      !rfidReader.PICC_ReadCardSerial()) {
    return;
  }

  String cardNumber = "";
  for (byte i = 0; i < rfidReader.uid.size; i++) {
    cardNumber += String(rfidReader.uid.uidByte[i], HEX);
  }

  String currentTime = getCurrentTime();
  String timeIn = "-", timeOut = "-", status = "Unknown";

  int userIndex = findUserIndex(cardNumber);

  if (userIndex != -1) { // Known RFID
    UserData &user = userList[userIndex];

    if (!user.isInside) { // Student entering
      user.lastTimeIn = currentTime;
      user.isInside = true;
      timeIn = user.lastTimeIn;
      status = "STUDENT ENTERED";
    } else { // Student exiting
      timeIn = user.lastTimeIn;
      timeOut = currentTime;
      user.isInside = false;
      status = "STUDENT EXITED";
    }

    Serial.print(serialCounter++);
    Serial.print(", ");
    Serial.print(cardNumber);
    Serial.print(", ");
    Serial.print(user.rollNumber);
    Serial.print(", ");
    Serial.print(user.name);
    Serial.print(", ");
    Serial.print(timeIn);
    Serial.print(", ");
    Serial.print(timeOut);
    Serial.print(", ");
    Serial.println(status);
  } else { // Unknown RFID
    Serial.print(serialCounter++);
    Serial.print(", ");
    Serial.print(cardNumber);
    Serial.print(", ");
    Serial.print("Unknown");
    Serial.print(", ");
    Serial.print("Unknown");
    Serial.print(", ");
    Serial.print(currentTime);
    Serial.print(", ");
    Serial.print("-");
    Serial.print(", ");
    Serial.println("UNKNOWN CARD ENTERED");
  }

  delay(1000);
  rfidReader.PICC_HaltA();
  rfidReader.PCD_StopCrypto1();
}
