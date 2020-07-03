#include <LiquidCrystal.h> //adicionando biblioteca para LCD
#include <DallasTemperature.h> //adicionando biblioteca para DS18B20
#include <OneWire.h> //adicionando biblioteca para capturar os dados do sensor

#define ONE_WIRE_BUS 8 //portas dos pinos de sinal dos DS18B20's

OneWire oneWireObjeto(ONE_WIRE_BUS); //define uma instÃ¢ncia do OneWire para comunicaÃ§Ã£o com o sensor
DallasTemperature sensorDS18B20(&oneWireObjeto);
 
DeviceAddress T1 = {0x28, 0xFF, 0x73, 0x08, 0xB3, 0x17, 0x01, 0xF3};
DeviceAddress T2 = {0x28, 0xFF, 0x63, 0x58, 0xB1, 0x17, 0x04, 0x27};
DeviceAddress T3 = {0x28, 0xFF, 0x92, 0x06, 0xB3, 0x17, 0x01, 0xF1};

LiquidCrystal lcd(12, 11, 7, 6, 5, 4); //Definindo os pinos que serÃ£o usados no LCD

void setup() {
    
  Serial.begin(9600); //Inicia o Monitor Serial
  sensorDS18B20.begin(); //Inicia comunicaÃ§Ã£o com os sensores
  lcd.begin(16, 2); //Inicia o LCD
  lcd.clear();
  lcd.setCursor(0,0);
  lcd.print("Iniciando...");
  delay(1000);
  }

void loop() {
 
  sensorDS18B20.requestTemperatures(); //Requisitando a temperatura do sensor

  Serial.print(sensorDS18B20.getTempC(T1));
  Serial.print(" , ");
  Serial.print(sensorDS18B20.getTempC(T2));
  Serial.print(" , ");
  Serial.println(sensorDS18B20.getTempC(T3));
  
  float temp1=sensorDS18B20.getTempC(T1);
  float temp2=sensorDS18B20.getTempC(T2);
  float temp3=sensorDS18B20.getTempC(T3);
 
  lcd.clear();
  lcd.setCursor(0, 0);
  lcd.print("1:");
  lcd.setCursor(2, 0);
  lcd.print(temp1);
  lcd.setCursor(6,0);
  //lcd.write(223);
  lcd.print("C ");
  
  lcd.setCursor(8, 0);
  lcd.print("2:");
  lcd.setCursor(10,0);
  lcd.print(temp2);
  lcd.setCursor(14,0);
  //lcd.write(223);
  lcd.print("C ");
  
  lcd.setCursor(0, 1);
  lcd.print("3:");
  lcd.setCursor(2, 1);
  lcd.print(temp3);
  lcd.setCursor(6,1);
  lcd.print("C ");

  delay(100);
 
}
