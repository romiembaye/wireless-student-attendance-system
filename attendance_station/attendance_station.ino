#include <SPI.h>
#include <MFRC522.h>
#include <LiquidCrystal_I2C.h>
#include <PubSubClient.h>
#include <ESP8266WiFi.h>

void setUpPinsAndLCD();
void reconnect();
void notifyProcessing();
void notifySuccess();
void notifyError();
void scanCards();
void lcdDisplay(String topRow, String bottomRow);
void mqttMessage(char* topic, byte* payload, unsigned int length);

#define BLUELED         D0
#define REDLED          D1
#define GREENLED        D2
#define BUZZER          D8
#define RFIDRESET       D3
#define RFIDSS          D4
#define LCDCLOCK        3
#define LCDDATA         1
#define MQTTIP          "*.*.*.*"
#define WIFINAME        "*******"
#define WIFIPASS        "*******"
#define PORT            1883

String clientName = "ESP8266-Attendance";
String idCard = "";
boolean systemActive = false;
boolean continueScanning = true;

WiFiClient wifiClient;
MFRC522 mfrc522(RFIDSS, RFIDRESET);
LiquidCrystal_I2C lcd(0x27, 16, 2);
PubSubClient client(MQTTIP, PORT, mqttMessage, wifiClient);

void setup() {
  setUpPinsAndLCD();
  SPI.begin();
  mfrc522.PCD_Init();
  WiFi.mode(WIFI_STA);
  WiFi.begin(WIFINAME, WIFIPASS);
  reconnect();
  delay(500);
}

void loop() {
  if (!client.connected() && WiFi.status() == 3) {
    reconnect();
  }
  client.loop();
  delay(10);
  scanCards();
}

void scanCards(){
  if (systemActive && continueScanning){
    idCard = "";
    lcdDisplay("  Scan your ID  ", "                ");
    if ( mfrc522.PICC_IsNewCardPresent()){
      if ( mfrc522.PICC_ReadCardSerial()){
        for (byte i = 0; i < mfrc522.uid.size; i++) {
          idCard += String(mfrc522.uid.uidByte[i], HEX);
        }
        client.publish("ATTENDANCE", idCard.c_str());
        notifyProcessing();
        mfrc522.PICC_HaltA();
        continueScanning = false;
      }
    }
  } else if (systemActive && !continueScanning) {
    lcdDisplay("     Status     ", "   Waiting...   ");
  } else{
    lcdDisplay("     Status     ", "    Inactive    ");
  }
}

void notifyProcessing(){
  digitalWrite(BUZZER, HIGH);
  delay(100);
  digitalWrite(BLUELED, HIGH);
  digitalWrite(BUZZER, LOW);
  delay(1000);
  digitalWrite(BLUELED, LOW);
}

void notifySuccess(){
  digitalWrite(GREENLED, HIGH);
  delay(1000);
  digitalWrite(GREENLED, LOW);
}

void notifyError(){
  digitalWrite(REDLED, HIGH);
  delay(1000);
  digitalWrite(REDLED, LOW);
}

void setUpPinsAndLCD(){
  pinMode(BLUELED, OUTPUT);
  digitalWrite(BLUELED, LOW);
  pinMode(REDLED, OUTPUT);
  digitalWrite(REDLED, LOW);
  pinMode(GREENLED, OUTPUT);
  digitalWrite(GREENLED, LOW);
  pinMode(BUZZER, OUTPUT);
  digitalWrite(BUZZER, LOW);
  pinMode(3, FUNCTION_3);
  pinMode(3, OUTPUT);
  pinMode(1, FUNCTION_3);
  pinMode(1, OUTPUT);
  lcd.begin(LCDDATA, LCDCLOCK);
  lcd.backlight();
  lcd.clear();
}

void lcdDisplay(String topRow, String bottomRow){
  lcd.setCursor(0, 0);
  lcd.print(topRow);
  lcd.setCursor(0, 1);
  lcd.print(bottomRow);
}

void mqttMessage(char* topic, byte* payload, unsigned int length) {
  if (payload[0] == 'S'){
      lcdDisplay("  Scan your ID  ", "Scanned " + idCard);
      notifySuccess();
  } else if (payload[0] == 'E'){
      lcdDisplay("  Scan your ID  ", "Scanning Failed");
      notifyError();
  } else if (payload[0] == 'A'){
      lcdDisplay("  Scan your ID  ", "Already Scanned");
      notifyError();
  } else if (payload[0] == '1'){
      systemActive = true;
      lcdDisplay("     Status     ", "     Active     ");
  } else if (payload[0] == '0'){
      systemActive = false;
  }
  continueScanning = true;
  delay(1000);
}

void reconnect() {
  if (WiFi.status() != WL_CONNECTED) {
    lcdDisplay("  Establishing  ", "WIFI  Connection");
    while (WiFi.status() != WL_CONNECTED) {
      delay(500);
    }
  }
  if (WiFi.status() == WL_CONNECTED) {
    while (!client.connected()) {
      if (client.connect((char*) clientName.c_str())) {
        client.subscribe("STATION");
        client.publish("ATTENDANCE", "1");
      } else {
        delay(500);
      }
    }
  }
}
