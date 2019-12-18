package com.test.utils;

import com.test.entity.User;
import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

import javax.servlet.http.HttpServletResponse;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

@Data
@SuppressWarnings("all")
public class ExcelExportUtil<T> {

  private int rowIndex;
  private int styleIndex;
  private String templatePath;
  private Class clazz;
  private Field fields[];

  public ExcelExportUtil(Class clazz,int rowIndex,int styleIndex){
    this.clazz = clazz;
    this.rowIndex = rowIndex;
    this.styleIndex = styleIndex;
    fields = clazz.getDeclaredFields(); //反射
  }


  /**基于注解的导出  扫描注解*/
  public  void export(HttpServletResponse response, InputStream is, List<T> objs,String fileName) throws IOException, IllegalAccessException {
    XSSFWorkbook workbook = new XSSFWorkbook(is);
    Sheet sheet = workbook.getSheetAt(0);
    CellStyle[] styles = getTemplateStyles(sheet.getRow(styleIndex));
    //原子性数据
    AtomicInteger dataAi = new AtomicInteger(rowIndex);

    for(T t : objs){
      Row row = sheet.createRow(dataAi.getAndIncrement());
      for(int i = 0; i<styles.length; i ++){
        Cell cell = row.createCell(i);
        cell.setCellStyle(styles[i]);

        /**********************标记：讲完反射回头看***************************/
        for(Field field : fields){
          if(field.isAnnotationPresent(ExcelAttribute.class)){
            field.setAccessible(true); //暴力获得
            ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
            if(i == ea.sort()){
              cell.setCellValue(field.get(t).toString());
            }
          }
        }
      }
    }
    String filename = URLEncoder.encode(fileName, "UTF-8");
    response.setContentType("application/octet-stream");
    response.setHeader("content-disposition","attachment;filename=" + new String(filename.getBytes("ISO-8859-1")));
    response.setHeader("filename",filename);
    workbook.write(response.getOutputStream());
  }


  public CellStyle[] getTemplateStyles(Row row){
    CellStyle [] styles = new CellStyle[row.getLastCellNum()];
    for(int i = 0; i < row.getLastCellNum() ; i ++){
      styles[i] = row.getCell(i).getCellStyle();
    }
    return styles;
  }


  public static <T> void export(HttpServletResponse response,List<T> objs,Class clazz,String fileName,String pathName) throws IOException, IllegalAccessException {
    ExcelExportUtil<T> excelExportUtil = new ExcelExportUtil(clazz,1,0);
    Resource resource = new ClassPathResource(pathName);
    FileInputStream in = new FileInputStream(resource.getFile());
    excelExportUtil.export(response,in, objs,fileName);
  }


}
