# farm_system_union(IoTogether) / 12월 2주차 기술보고서
##  - (산업시스템공학과 2022112386 김호준)

## 새로운 시각화 도구

저번 주까지는 LCD와 텔레그램을 활용하여 문자를 이용한 데이터 출력을 시각화할 수 있었다. 그저 공지용으로 해당 시스템을 이용한다면 충분한 시스템으로 보인다.

하지만 본 조가 구현하고자 하는 시스템은 학업에 집중하고 있는 학생들을 위한 시스템이다보니, 데이터를 보기 편한 형태로 시각화하는 것이 낫겠다고 필자는 판단한다.

따라서 '아두이노를 통해 데이터를 수집하고, 그 정보를 '엑셀'에 저장하면 괜찮지 않을까'라는 생각이 들기도 하였다.


## 엑셀의 기능들이자 장점

엑셀의 가장 큰 특장점은 역시나 데이터 정리하는 데에 있어서 가장 편한 도구다. 행과 열마다 Name을 붙여주고 각각의 이름에 해당하는 주소에 데이터 값들을 넣어 나중에도 편하게 자료를 찾을 수 있다는 점이 큰 특징이다.

즉, 우리가 원하는 자료의 시각화와 더불어 '데이터 저장'도 함께 수행할 수 있다는 점이다. 이렇게 데이터를 수집하여, 언제 습도나 소음이 큰지에 대한 분석하는 데에 큰 도움을 줄 수 있다.

엑셀의 장점은 여기서 멈추지 않고, 그래프를 그려주는 데에 큰 이점이 있다. 특정 행렬 집단을 동시 선택하여 꺽은선, 박스, 막대 플롯 등 다양한 형태의 시각화를 엑셀에서 자체 기능으로 지원한다.

엑셀이라는 '하나의' 프로그램에서 필자가 원하는 기능을 모두 제공해준다는 점에서, 다른 프로그램을 사용할 필요가 없는, 오히려 다기능이기에 역설적으로 데이터 수집 및 정리가 간단하게 만들어 주는 좋은 프로그램이다.

마지막으로, 엑셀은 마이크로소프트社 에서 제공한 프로그램이기에, 자사 계정만 있으면 '원드라이브' 클라우드 서비스도 또한 활용할 수 있다.

링크만 가지고 있다면 해당 데이터를 볼 수 있는 권한이 생기기에 얼마나 편리한 시스템인가??

<div align="center">
 <img width="150"  src="https://github.com/user-attachments/assets/707f6d1c-c8a1-4749-aa0e-9714cc88e311" />
 <img width="250"  src="https://github.com/user-attachments/assets/fff20d17-56bc-4617-9fa9-c47a5fcd135a" />
</div>


## 소프트웨어 세팅

아두이노에서는 새로 추가할 라이브러리나 함수는 그다지 않고, 대신에 새로운 프로그램을 다운로드 받아야 한다.

PLX-DAQ라는 프로그램으로 엑셀의 매크로 시스템을 이용하여 자동으로 데이터를 정리해주는데, 아두이노와 잘 연동되도록 설정(예를 들어, 시리얼 통신 포트와 속도)을 건들 수 있게 프로그래밍이 되어있다보니,

필자가 직접 조작한 전체의 80% 영역은 해당 프로그램에서 이루어졌다고 해다고 관언이 아니다. 프로그램 설치 방법과 그 내부를 살펴보자.

### PLX-DAQ 설치

<div align="center">
 <img width="800" src="https://github.com/user-attachments/assets/4cf61fb3-c474-43af-a53b-b8237876826c" />
 
 < 사이트 주소 : https://forum.arduino.cc/t/plx-daq-version-2-now-with-64-bit-support-and-further-new-features/420628 >
</div>
  
아두이노 포럼에서 직접 시스템을 제공하고 있음을 확인하여, 가장 최신 버전인 2.11로 선택하였다.

<div align="center">
 <img width="800" src="https://github.com/user-attachments/assets/b165627d-cceb-4270-b43f-8433845c57a2" />
</div>

버전을 선택하고 다시 내부사이트로 들어가서, 스크롤을 조금만 내리면 .zip 파일의 형태로 소프트웨어를 다운로르받을 수 있다.

그 후 다운로드 된 파일의 압축을 풀어주면 네 개의 파일을 확인할 수 있는데, 이 중 확장자명이 .xlsm로 끝나는 파일을 "엑셀"로 연결 프로그램으로 설정하여 열어주면 소프트웨어 화면을 확인할 수 있다.

<div align="center">
 <img width="800" src="https://github.com/user-attachments/assets/6d52f26c-547a-4999-a930-dca2787f248e" />

 <img width="800" src="https://github.com/user-attachments/assets/c6503434-490b-4c8e-a63f-f4d05a07e489" />
 
 < 엑셀 내부에서의 모습 / PLX DAQ로 시작하는 새로운 팝업창을 볼 수 있다. >
</div>

팝업창에서 아두이노 시리얼 환경에 맞게 Port번호와 통신속도(Baud)를 적절히 설정한다. 참고로 필자는 포트번호를 "7", 통신속도는 "9600"으로 설정하였다. PLX-DAQ는 115200 속도는 통신이 잘 안된다고 하는데, 혹시 모르니(안정성의 문제) 필자도 해당 주의에 따르도록 하겠다.

<div align="center">
 <img width="400" src="https://github.com/user-attachments/assets/7c95322a-7c61-433c-964a-08ca082fdcfd" />
</div>

이대로 나중에 아두이노 회로를 PC와 연결하고, 통신이 시작된 상태에서 팝업창의 "Connect"를 누르면 자동으로 데이터를 저장할 수 있다.

### 아두이노 내부 세팅

아두이노도 코드 몇 줄을 더 추가해줘야 한다. 다만, 선술했듯이 따로 라이브러리를 반영할 필요 없다. 추가해야 하는 코드들을 보자.

```
 // 엑셀에 표시할 내용을 제한하는 변수
 int count = 0;

 void setup() {

    // 통신속도는 PLX-DAQ의 설정에 맞게 설정한다.
    Serial.begin(9600);

    // 데이터시트 초기화
    Serial.println("CLEARDATA"); // 제목 줄(두 번째)을 제외한 모든 내용을 지운다.

    // 표의 제목을 짓기 위한 코드, () 괄호 안에 "LABEL"로 시작한다.
    Serial.println("LABEL,Temp,Humi,Count");
}

 void loop() {

    // 온습도 센서에서 수집한 습도, 온도 정보
    float humi = dht.readHumidity();
    float temp = dht.readTemperature();

    // 루프마다 count 변수의 값이 1씩 늘어난다.
    count += 1;

    // 수집한 자료를 엑셀에 저장하는 코드,  첫 번째 출력에서 "DATA,"로 시작한다.
    // 옆으로 자료를 작성하려면 쉼표(",")로 구분한다.
    // 아래쪽으로 작성하려면 println()을 이용한다.
    Serial.print("DATA,");      
    Serial.print(temp);
    Serial.print(",");
    Serial.print(humi);
    Serial.print(",");
    Serial.println(count);

    // if문을 통해 count값이 30이 되면 다시 데이터를 초기화한다.(데이터시트를 지운다)
    // 파일의 크기를 제안하여 원활한 데이터 수집을 만들어낸다.
    // 두 개의 출력문을 Serial.println("DATA"); 로 줄여도 된다.
    if (count >= 30) {
        Serial.println("CLEARSHEET");
        Serial.println("LABEL,Temp,Humi,Count");
        count = 0;
    }
 }
```

해당 구문을 이전 주에 작성한 코드에 추가하여보자. 대신 LCD는 더이상 시각화의 도구를 사용하지 않을 예정이므로, LCD 출력 구문들은 모두 지워서 반영하였다.

```
#include <UniversalTelegramBot.h>
#include <WiFiClientSecure.h>
#include <ESP8266WiFi.h>
#include <DHT.h>

char message[100];

#define DHTPIN D8
#define DHTTYPE DHT11

DHT dht(DHTPIN, DHTTYPE);

#define WIFI_SSID ""//와이파이 이름
#define WIFI_PASSWORD ""//와이파이 비밀번호

#define BOT_TOKEN "8140194638:AAGOpEGxQ9Qf3nxU6sLhEFPsJLk23sPn7Jw"  //bot token
#define CHAT_ID "5561347315"     //Chat ID

WiFiClientSecure client;
UniversalTelegramBot bot(BOT_TOKEN, client);

int count = 0;

void setup() {
 Serial.begin(9600);

 Serial.println("");
 Serial.println("와이파이에 연결 중...");
 WiFi.begin(WIFI_SSID, WIFI_PASSWORD);
 while (WiFi.status() != WL_CONNECTED){
   Serial.print(".");
   delay(1000);
 }
 Serial.println("");
 Serial.println("와이파이 연결 완료.");
 Serial.println();

 configTime(0, 0, "pool.ntp.org");
 time_t now = time(nullptr);
 Serial.println("시간 동기화 중...");
 while (now < 24 + 3600){
   now = time(nullptr);
   Serial.print(".");
   delay(1000);
 }
 Serial.println("");
 Serial.println("시간 동기화 완료.");
 Serial.println("");

 Serial.println("CLEARDATA");
 Serial.println("LABEL,Temp,Humi,Count");

 dht.begin();
 
}

void loop() {
 float humi = dht.readHumidity();
 float temp = dht.readTemperature();
 
 count += 1;

 Serial.print("DATA,");
 Serial.print(temp);
 Serial.print(",");
 Serial.print(humi);
 Serial.print(",");
 Serial.println(count);
 

 if (temp >= 30 || humi >= 70){
   Serial.println("공부하기에 어렵습니다.");

   if (WiFi.status() != WL_CONNECTED){
     Serial.println("WiFi에 연결이 되어있지 않아 텔레그램으로 메시지를 전송할 수 없습니다.");
   }
   else {
     sprintf(message, "공부하기에 어렵습니다. : %f C / %f %", temp, humi);
     Serial.println("텔레그램을 통해 문자를 보냅니다.");
     X509List cert(TELEGRAM_CERTIFICATE_ROOT);
     client.setTrustAnchors(&cert);
     bot.sendMessage(CHAT_ID, message, "");
   }

   for (int i = 0; i < 10; i++){
     delay(1000);
   }
 }
 if (count >= 30) {
 Serial.println("CLEARSHEET");
 Serial.println("LABEL,Temp,Humi,Count");
 count = 0;
 }
 delay(500);
}
```

### 실행 모습

<div align="center">
 <img width="800" src="https://github.com/user-attachments/assets/cea07028-b852-40ba-a06a-cc0187a58c79" />

 <img width="800" src="https://github.com/user-attachments/assets/3180c9a4-70f2-46d0-be52-5aefb4c3f156" />
</div>

각각의 열에 맞춰 자료를 반영하는 모습을 확인할 수 있었다. 특히 필자는 엑셀 파일의 용량 걱정을 우려하여 행의 개수가 30을 넘기지 않도록 설정하였는데, count 제한 수를 늘려 더 많은 데이터를 수집할 수 있도록 설정할 수 있다.


### 원드라이브의 이용

엑셀은 사용자가 만일 원드라이브를 소유하고 있으면, 간단한 설정을 하여 파일을 클라우드 상으로 저장할 수 있도록 한다. 즉, 인터넷을 이용한 접근이 가능하므로 링크를 보유한 사용자들은 쉽게 자료를 열람할 수 있다는 것이다.

순서는 

 - 먼저 원드라이브에 해당 파일을 업로드(자동 저장 켜기)하고, 

<div align="left">
 <img width="450"  src="https://github.com/user-attachments/assets/8c5014e2-2547-4d16-929d-6366dc6daa1a" />
</div>

 - '공유' 버튼을 눌러 '링크 복사'한다면,

<div align="left">
 <img width="450" src="https://github.com/user-attachments/assets/dfc48084-8acf-4d2a-ae26-8ba6c5bc611f" />
</div>

 - 클립보드로 링크가 저장된다.

<div align="left">
 <img width="600" src="https://github.com/user-attachments/assets/0bd84ddc-0492-454c-86f0-34dceb8ac8f5" />
</div>

<태블릿(모바일, 안드로이드)>

<div align="center">
 <img width="800" src="https://github.com/user-attachments/assets/d0194d3f-7636-40c5-95b1-ed3a1dd02689" />
</div>

<웹 환경>

<div align="center">
 <img width="800" src="https://github.com/user-attachments/assets/5bc9412c-506d-4eea-9ca5-761aeadf6057" />
</div>

다행히 두 환경에서 자료 수집의 모습을 잘 보여주었다.

하지만 원드라이브는 아직 모바일 친화적인지 않은 것인지, 태블릿 환경은 처음에는 잘 데이터를 표시해주었다가 나중에는 기능을 멈춰버리고, 휴대폰으로는 아예 데이터 수집조차도 어려운 모습을 보였다.

그나마 대안은 모바일도 웹에서 데이터를 잘 받는 모습을 보여 원드라이브 매크로 엑셀 내용은 일단 '앱'이 아닌 '웹'으로 확인하는 방법이다.

여기서 멈추지 않고, 엑셀의 자체기능인 그래프 표시도 이용해보자.

< PC 엑셀 >

<div align="center">
 <img width="800" src="https://github.com/user-attachments/assets/d4eddb13-fdce-4078-8e91-57d9c44c4d01" />
</div>

< PC 웹 >

<div align="center">
 <img width="800" src="https://github.com/user-attachments/assets/942fdcae-a80d-47e7-b8c6-9fc54d9650da" />
</div>

< 태블릿 웹 >

<div align="center">
 <img width="800" src="https://github.com/user-attachments/assets/099400db-3126-4d6c-a338-c4a4db13369a" />
</div>

웹 환경에서 그래프까지 모두 잘 보여준다.

설정을 잘 건드려, 링크로 초대받은 사용자는 내부 자료를 수정할 수 없도록 관리한다면 좋은 시각화 도구로 사용할 수 있다고 판단한다. 
