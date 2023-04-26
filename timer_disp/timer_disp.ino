#include <SPI.h>
#include <Wire.h>
#include <Adafruit_GFX.h>
#include <Adafruit_SSD1306.h>

Adafruit_SSD1306 display(-1);

String x;
String N;
int H;
int M;
int S;

void setup() 
{  
   Serial.begin(115200);
   Serial.setTimeout(2);
   
   read_serialdata();
   
}

void read_serialdata(void) // to read and parse the incoming data from the py software.
{
  while (!Serial.available());
x = Serial.readString();

int commaIndex = x.indexOf(',');
int secondCommaIndex = x.indexOf(',', commaIndex + 1);
int thirdCommaIndex = x.indexOf(',', secondCommaIndex +1);

String Name = x.substring(0, commaIndex);
String Hours = x.substring(commaIndex + 1, secondCommaIndex);
String Minutes = x.substring(secondCommaIndex + 1, thirdCommaIndex);
String Seconds = x.substring(thirdCommaIndex + 1);

Serial.print(Name);
Serial.print(Hours);
Serial.print(Minutes);
Serial.print(Seconds);
Serial.flush();
N = Name;
H = Hours.toInt();
M = Minutes.toInt();
S = Seconds.toInt();


}


String Time_disp(int hrs, int mins, int sec) //Count down timer display
{
  
}

void loop() 
{
      display.begin(SSD1306_SWITCHCAPVCC, 0x3C);  
  if(!display.begin(SSD1306_SWITCHCAPVCC, 0x3C)) 
    {
      Serial.println(F("SSD1306 allocation failed"));
      for(;;);
    }
  delay(2000);
  display.clearDisplay();
  display.setTextColor(WHITE);
    // Clear the buffer.
  display.clearDisplay();

  // Scroll part of the screen
  display.setTextSize(1);
  display.setTextColor(WHITE);
  display.setCursor(0,0);
  display.println("Name of the patient: ");
  display.println(N);
  display.display();
//  display.startscrollright(0x01, 0x01);

  display.setTextSize(1);
  display.setTextColor(BLACK,WHITE);
  display.setCursor(48,16);
  display.println(" Time ");
  display.display();


    display.setTextSize(1);
    display.setTextColor(WHITE);
    display.setCursor(48,24);
    display.println(Time_disp(H,M,S));
    display.display();
}
