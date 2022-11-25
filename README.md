# java-excelRead
excel Read  java source

엑셀파일 업로드하고 데이터 처리
***

```java
// service 호출로직
public ResponseCode excelUpload(NonPayment param, MultipartRequest multipartReq) {
  MultipartFile uploadFile = multipartReq.getFile("excelFile");
  // excel file Upload
  File file = new File(path + uploadFile);
  // excel file READ
  List<Map<String, String>> excelContent = ExcelData(file);
  
  // 이후 excelContent 데이터 DB 처리
```

```java
private List<Map<String, String>> ExcelData(File FILE) {
  ExcelReadOption excelReadOption = new ExcelReadOption();
  excelReadOption.setFilePath(FILE.getAbsolutePath());
  
  // 데이터 col 갯수 (임의 강제지정)
  excelReadOption.setOutputColumns("A","B","C","D","E","F","G","H","I","J","K","L");
  excelReadOption.setStartRow(2);

  // ex) Key : A , Value : CODE_VALUE
  // ex) Key : F , Value : 10000
  List<Map<String, String>> excelContent = ExcelRead.read(excelReadOption);

  return excelContent;
}
```
***

### 운영 중 이슈 발생.
1. 엑셀파일 셀수가 일정하게 작성하지 않는 경우.
2. 엑셀 편집시 행삭제가 올바르게 수행되지 않는 경우.
  
```java
  // 1번 이슈 조치
  /**
  * getPhysicalNumberOfCells()
  * >> 빈 값은 체크 안함 : 빈 셀들은 체크 안함.
  */
  int numOfCells = row.getPhysicalNumberOfCells();

  /**
  * getLastCellNum()
  * >> 마지막 cell 이 공백일 경우도 원하는 컬럼수가 나오지 않음.
  */ 
  int numOfCells = row.getLastCellNum();

  /**
  * excelReadOption.setOutputColumns("A","B","C","D","E","F","G","H","I","J","K","L");
  * >> 임의 강제지정한 셀수로 고정
  */ 
  int numOfCells = excelReadOption.getOutputColumns().size();
  ```

```java
// 2번 이슈 조치
/* 
 * Excel 파일 행삭제가 올바르게 수행되지 않았을 경우  row 가 null로 잡히지 않고 통과되어
 * 빈 Row 값들이 삽입되는 경우가 있음.
 * 
 * @DESCRPTION : 첫번재 열이 공백일경우 데이터 없는 것으로 간주하고 종료처리
 */
if( cellName.equals("A") && ExcelCellRef.getValue(cell).equals("")) {
  break LoopRow;
}
```



엑셀관련소스 활용블로그 : <https://m.blog.naver.com/PostView.naver?isHttpsRedirect=true&blogId=0oooox&logNo=220343916758>
