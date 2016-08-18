# ExcelTableBuilder

* 엑셀 파일을 테이블 데이터를 바이너리로 빌드하는 툴입니다.
* C# 기반을 사용하는 클라이언트(ex. Unity3D) 및 서버에서 공통된 데이터 바이너리를 간단히 공유하고 작업할 수 있도록 도와주는 툴입니다.
* 간단한 방식의 엑셀 기반 규격을 가지고 동작합니다. 엑셀이 설치된 운영체제에서만 동작합니다.

#####빌드 결과로 생성되는 결과물은 총 3가지 입니다.
1. 최종 결과 바이너리 파일(기본 파일명: excel_table_bin.bytes)
 * 생성된 소스 코드에 로드될 파일입니다.
 * 버전 컨트롤이 필요 없습니다.
2. 생성된 소스 코드 (GeneratedTableSource.cs)
 * 테이블 데이터를 사용할 소스에서 포함해서 사용하면 됩니다.
 * 버전 컨트롤이 필요합니다.
 * 엑셀 파일로 정의된 내용대로 생성된 클래스 소스 등이 들어갑니다.
3. 프리파싱 파일
 * 빌드 툴을 위해 존재하는 프리파싱 파일입니다.
 * 파싱 정보를 캐시해서 재 빌드시 변경된 파일만 재 빌드하도록 도와줍니다.
 * 버전 컨트롤이 필요합니다.

#####간략한 설명
 * 툴에서는 엑셀 데이터를 그대로 바이너리화 시키는 것이 아닌 C# 클래스 타입의 데이터를 생성한 후, 리플렉션을 통해 생성된 클래스 자체를 직렬화 해서 저장합니다.
 * 읽어들일 때는 속도 등을 위해 전체 데이터를 파싱하는 것이 아니라 해당하는 테이블의 행 단위로 파싱하고 캐싱합니다.
 * 초기 가동시에는 단순히 최종 바이너리(bytes) 파일을 로드하는 것 이외에 역직렬화를 위한 시간을 사용하지 않습니다.
 * 읽었던 테이블의 행은 캐시해두므로, 다시 접근할 때는 단순히 딕셔너리에 접근하는 것과 같습니다.
 * 직렬화 관련해서 TableBinConverter 라는 클래스를 이용해 클래스 프로퍼티를 바이너리로 직렬화 및 역직렬화를 합니다.
 구글 프로토버프를 사용하는 것이 더 빠르지만, 외부 라이브러리 의존성을 줄이고 간단히 사용하기 위해 변환기를 내장했습니다. 속도는... 프로토버프보다 조금 느립니다.

#####테이블 구성 및 사용법
 * 기본적인 구조는 Sample1.xlsx 파일을 참조해주세요.
 * !TABLE 로 테이블 시작 위치와 테이블 이름을 지정합니다.
 * !FIELD_NAME 으로 필드 이름 행을 지정하고, 필드 종료 부분에 !EOF 로 필드 종료를 마킹합니다.
 * !FIELD_TYPE 으로 필드의 타입을 지정합니다. 이 형식에 따라 생성되는 소스 코드의 C# 클래스의 데이터 타입이 지정됩니다.
 * !EOT 로 테이블 종료 위치를 지정합니다.

 * 지원되는 필드 타입
  + !KEY ; 키 타입. 문자열로 구성되어 있고, TableKeyEnum 으로 관리됩니다. 모든 키는 테이블을 떠나 유니크 해야 합니다. 다른 테이블에서 키는 참조 가능합니다.
  + !KEY_ENUM ; 데이터 테이블이 아닌 Enum 테이블에서 키를 가져와서 정의할 때 사용하는 키 입니다.
  + !ID ; 키를 값으로 사용(같은 테이블이나 다른 테이블에서 키 참조)하는 경우에 사용합니다.
  + !STRING ; 문자열(유니코드)
  + !INT ; int
  + !INT_LIST ; 셀 내에 1,2,3 과 같이 숫자를 여러개 입력할 때 사용합니다. List<int> 타입입니다.
  + !FLOAT ; float
  + !FLOAT_LIST ; List<float>
  + !BOOL ; 부울 형식. true/false, y/n, 1/0 을 지원합니다. 대소문자를 가리지 않습니다. true/y/1 이외엔 모두 false 입니다.
  + !DATETIME ; 날짜 형식. 엑셀 OAD 규격의 날짜 데이터를 double로 가져옵니다. DateTime.FromOADate 형식으로 값을 얻어 올 수 있습니다.
  + !VECTOR ; Vector3 형식
  + !ENUM_KEY ; Enum 형태의 키 타입 선언 (Enum 전용 테이블)
  + !ENUM_ID ; Enum 형태의 키를 컬럼 값으로 사용하는 경우
  + !ENUM_VALUE ; Enum 형태의 키의 값 (ENUM_KEY 와 함께 사용합니다)

#####사용 예제
<pre><code>var table_data_manager = TDB.TableDataManager.LoadBinaryFile("excel_table_bin.bytes");
TDB.TableSample1 table_sample1;
if (table_data_manager.TableSample1.TryGetValue((int)TDB.TableKeyEnum.B10003, out table_sample1) == false)
{
    throw new Exception($"KEY NOT FOUND");
}

foreach (var s in table_sample1.sample_string)
{
    Console.WriteLine(s);
}</code></pre>

#####생성된 소스 파일
* GeneratedTableSource.cs을 참조해주세요.
