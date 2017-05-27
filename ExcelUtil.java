package com.liuwei.test;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jef.tools.IOUtils;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ExcelUtil {
	
	/**
	 * 读取Excel的内容,如果一行的前两列为空则返回，不再读取下面的行
	 * @param file
	 * @param ignoreRows 忽略行数，如果有头设成1
	 * @param colArray 列字段的映射
	 * @return
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static List<Map<String, Object>> getData(File file, int ignoreRows,String []colArray)throws Exception {
		List<Map<String, Object>> list = new ArrayList<Map<String,Object>>();
		int rowSize = 0;
		BufferedInputStream in = new BufferedInputStream(new FileInputStream(file));
		// 打开HSSFWorkbook
		POIFSFileSystem fs = new POIFSFileSystem(in);
		HSSFWorkbook wb = new HSSFWorkbook(fs);
		HSSFCell cell = null;
		int colSize = colArray.length;
		boolean isOver = false;//如果当前行前两列数据均为空，不再读取下面数据
//		for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {
			try {
				HSSFSheet st = wb.getSheetAt(0);//只读取第一个sheet
				// 第一行为标题，不取
				for (int rowIndex = ignoreRows; rowIndex <= st.getLastRowNum(); rowIndex++) {
					HSSFRow row = st.getRow(rowIndex);
					if (row == null) {
						break;
					}
					if (row.getCell(0) == null && row.getCell(1)==null) {
						break;
					}
					int tempRowSize = row.getLastCellNum() + 1;
					if (tempRowSize > rowSize) {
						rowSize = tempRowSize;
					}
					Map<String, Object> r = new HashMap<String, Object>(rowSize);
					boolean cell0Empty = false;
					for (int columnIndex = 0; columnIndex <= row.getLastCellNum(); columnIndex++) {
						if (columnIndex >= colSize) {
							break;
						}
						String value = "";
						if ((cell = row.getCell(columnIndex)) != null) {
							switch (cell.getCellType()) {
								case HSSFCell.CELL_TYPE_STRING:
									value = cell.getStringCellValue();
									break;
								case HSSFCell.CELL_TYPE_NUMERIC:
									value = new DecimalFormat("0").format(cell.getNumericCellValue());
									break;
								case HSSFCell.CELL_TYPE_FORMULA:
									// 导入时如果为公式生成的数据则无值
									if (!cell.getStringCellValue().equals("")) {
										value = cell.getStringCellValue();
									} else {
										value = cell.getNumericCellValue() + "";
									}
									break;
								case HSSFCell.CELL_TYPE_BLANK:
									break;
								case HSSFCell.CELL_TYPE_ERROR:
									value = "";
									break;
								case HSSFCell.CELL_TYPE_BOOLEAN:
									value = (cell.getBooleanCellValue() == true ? "Y": "N");
									break;
								default:
									value = "";
							}
						}
						if (StringUtils.isBlank(value)) {
							if (columnIndex == 0) {
								cell0Empty = true;
							}else if (cell0Empty && columnIndex == 1) {
								isOver = true;
								break;
							}
						}
						String v = value.trim();
						r.put(colArray[columnIndex], v);
					}

					if (isOver) {
						break;
					}
					list.add(r);
				}
			} catch (Exception e) {
				throw e;
			}finally{
				in.close();
			}

//		}
		return list;
		
	}
	
	
	/**
	 * 创建excel文件
	 * @param file
	 * @param title
	 * @param content
	 * @return
	 */
	public static File makeExcel(File file, String []title,String[][] content){
		HSSFWorkbook workbook = new HSSFWorkbook();// 相当于JXL中的 WritableWorkbook

		HSSFSheet sheet = workbook.createSheet();
		if (title == null || content == null) {
			return null;
		}
		int colSize = title.length;
		HSSFRow one = sheet.createRow(0);
		for (int i = 0; i < colSize; i++) {
			HSSFCell c = one.createCell(i);
			c.setCellType(HSSFCell.CELL_TYPE_STRING);
			c.setCellValue(title[i]);
		}
		for (int i = 1; i < content.length+1; i++) {
			HSSFRow row = sheet.createRow(i);// 表示sheet第i+1行
			for (int j = 0; j < content[i-1].length; j++) {
				HSSFCell cell = row.createCell(j);// 该行第j+1列元素
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				cell.setCellValue(content[i-1][j]);
			}
		}
		// 新建一输出文件流
		// 把相应的Excel 工作簿存盘
		FileOutputStream fOut = null;
		try {
			if (!file.exists()) {
				file.createNewFile();
			}
			fOut = new FileOutputStream(file);
			workbook.write(fOut);
			fOut.flush();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(fOut);
		}
		return file;
	}
	

	public static void main(String[] args) throws Exception {
	    
	   }
}
