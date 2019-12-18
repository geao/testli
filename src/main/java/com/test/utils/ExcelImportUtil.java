package com.test.utils;

import com.test.entity.User;
import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@Data
@SuppressWarnings("all") //Excel导入工具类
public class ExcelImportUtil<T> {

  private Class clazz;
  private Field fields[];

  public ExcelImportUtil(Class clazz) {
    this.clazz = clazz;
    fields = clazz.getDeclaredFields();
  }



  /**
   * 基于注解读取excel
   */
  public List<T> readExcel(InputStream is, int rowIndex, int cellIndex) throws IOException, IllegalAccessException, InstantiationException, ParseException {

    List<T> list = new ArrayList<T>();

    T entity = null;


    XSSFWorkbook workbook = new XSSFWorkbook(is);
    Sheet sheet = workbook.getSheetAt(0);

    for (int rowNum = rowIndex; rowNum <= sheet.getLastRowNum(); rowNum++) {
      Row row = sheet.getRow(rowNum);
      entity = (T) clazz.newInstance();
      for (int j = cellIndex; j < row.getLastCellNum(); j++) {
        Cell cell = row.getCell(j);
        for (Field field : fields) {
          if (field.isAnnotationPresent(ExcelAttribute.class)) {
            field.setAccessible(true);
            ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
            if (j == ea.sort()) {
              field.set(entity, coverType(field,cell));
            }
          }
        }
      }
      list.add(entity);
    }

    return list;
  }

  //类型转换 将cell单元格格式转换为字段类型
  private Object coverType(Field field, Cell cell) throws ParseException {
    String fieldType = field.getType().getSimpleName();
    if ("String".equals(fieldType)) {
      return getValue(cell);
    }else if("Date".equals(fieldType)){
      return new SimpleDateFormat("yyyy年MM月dd日HH:mm:ss").parse(getValue(cell));
    }else if("int".equals(fieldType) || "Integer".equals(fieldType)){
      return Integer.parseInt(getValue(cell));
    }else if("double".equals(fieldType) || "Double".equals(fieldType)){
      return Double.parseDouble(getValue(cell));
    }else{
      return null;
    }
  }

  public String getValue(Cell cell) {
    if (cell == null) return "";
    switch (cell.getCellType()) {
      case STRING:
        return cell.getRichStringCellValue().getString().trim();
      case NUMERIC:
        if (DateUtil.isCellDateFormatted(cell)) {
          Date date = DateUtil.getJavaDate(cell.getNumericCellValue());
          return new SimpleDateFormat("yyyy年MM月dd日HH:mm:ss").format(date);
        } else {
          //防止数值编程科学计数法
          String strCell = "";
          Double num = cell.getNumericCellValue();
          BigDecimal bd = new BigDecimal(num.toString());
          if (bd != null) {
            strCell = bd.toPlainString();
          }
          //取出浮点型 自动加的.0
          if (strCell.endsWith(".0")) {
            strCell = strCell.substring(0, strCell.indexOf("."));
          }
          return strCell;
        }
      case BOOLEAN:
        return String.valueOf(cell.getBooleanCellValue());
      default:
        return "";
    }
  }


  public List<T> readExcel(Class clazz,InputStream is) throws InstantiationException, IllegalAccessException, ParseException, IOException {
    ExcelImportUtil<T> excelImportUtil = new ExcelImportUtil<>(clazz);
    List<T> list = excelImportUtil.readExcel(is, 1, 0);
    return list;
  }

}
