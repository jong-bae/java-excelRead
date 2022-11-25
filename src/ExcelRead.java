package com.jaseng.common.excel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelRead {
	public static List<Map<String, String>> read(ExcelReadOption excelReadOption) {

		//엑셀파일 확장자에 따라 Workbook 객체생성
		Workbook wb = ExcelFileType.getWorkbook(excelReadOption.getFilePath());
		
		/**
		 * getSheetAt(0) : 엑셀 파일에서 첫번째 시트 로드
		 * getSheet("시트명") : 시트명을 지정하여 해당 시트만 로드 
		 */
		//Sheet sheet = wb.getSheetAt(0);
		Sheet sheet = wb.getSheet("sheet1");
		
		//System.out.println("Sheet 이름: "+ wb.getSheetName(0));
		//System.out.println("데이터가 있는 Sheet의 수 :" + wb.getNumberOfSheets());
		
		/**
		 * sheet에서 유효한(데이터가 있는) 행의 개수를 가져온다.
		 */
		int numOfRows = sheet.getPhysicalNumberOfRows();
		int numOfCells = excelReadOption.getOutputColumns().size();

		Row row = null;
		Cell cell = null;

		String cellName = "";
		/**
		 * cell 값을 저장할 맵 객체
		 */
		Map<String, String> map = null;
		/*
		 * Row에 담을 List 변수
		 */
		List<Map<String, String>> result = new ArrayList<Map<String, String>>();
		
		/**
		 * 위에서 구한 numOfRows 크기만큼 반복조회
		 */
		LoopRow:
		for(int rowIndex = excelReadOption.getStartRow() - 1; rowIndex < numOfRows; rowIndex++) {
			// row 데이터 
			row = sheet.getRow(rowIndex);

			if(row != null) {
				/**
				 * 
				 * 아래와 같은 이유로 지정한 컬럼수를 가져오게끔 위에서 처리함 
				 * (excelReadOption.getOutputColumns().size();) 
				 * 
				 * 
				 * >> 빈 값은 체크 안함 : 빈 셀들은 체크 안함.
				 * int numOfCells = row.getPhysicalNumberOfCells();
				 * 
				 * >> 마지막 cell 이 공백일 경우도 원하는 컬럼수가 나오지 않음.
				 * int numOfCells = row.getLastCellNum();
				 */
			
				map = new HashMap<String, String>();
				/*
				 * cell의 수 만큼 반복한다.
				 */
				for(int cellIndex = 0; cellIndex < numOfCells; cellIndex++) {

					cell = row.getCell(cellIndex);
					cellName = ExcelCellRef.getName(cell, cellIndex);
					
					/*
					 * 추출 대상 컬럼인지 확인한다
					 * 추출 대상 컬럼이 아니라면,
					 * for로 다시 올라간다
					 */
					if( !excelReadOption.getOutputColumns().contains(cellName) ) {
					    continue;
					}
					
					/* 
					 * Excel 파일 행삭제로 삭제했을 경우  row 가 null로 잡히지 않고 통과되어
					 * 빈 Row 값들이 삽입되는 경우가 있음.
					 * 
					 * @UPDATE : 220407 - jb 
					 * @DESCRPTION : 첫번재 열이 공백일경우 데이터 없는 것으로 간주하고 종료처리
					 */
					if( cellName.equals("A") && ExcelCellRef.getValue(cell).equals("")) {
						break LoopRow;
					}
					/*
					 * map객체의 Cell의 이름을 키(Key)로 데이터를 담는다.
					 * put("A" 열, "값");
					 */
					map.put(cellName, ExcelCellRef.getValue(cell));
				}

				result.add(map);
			}
		}
		return result;
	}
	
}
