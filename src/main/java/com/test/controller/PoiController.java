package com.test.controller;

import com.test.entity.EmployeeReportResult;
import com.test.entity.User;
import com.test.utils.ExcelExportUtil;
import com.test.utils.ExcelImportUtil;
import org.apache.el.parser.ParseException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

@RestController
@RequestMapping("/excel")
@CrossOrigin
@SuppressWarnings("all")
public class PoiController {

  @GetMapping("/import")
  public void importExcel() throws IOException, ParseException, java.text.ParseException {
          File file = new File("D:\\user.xlsx");
    InputStream in = new FileInputStream(file);

    //创建workbook excel大对象
    Workbook workbook = WorkbookFactory.create(in);

    //获得第一个工作簿对象
    Sheet sheet = workbook.getSheetAt(0);

    //声明list集合来存储从Excel表中获得数据
    List<User> users = new ArrayList<>();

    //sheet.getLastRowNum(); //获得最后一行
    //从1开始是因为第一个行的索引是0,并且它是标题头,不需要存入数据库
    for(int i = 1 ; i <= sheet.getLastRowNum(); i ++ ){
      //获取行对象
      Row row = sheet.getRow(i);

      //声明对象

      //获取单元格
      //赋值给User实体
      User user = User.builder()
            .id(getCellValue(row.getCell(0)).toString())
            .name(getCellValue(row.getCell(1)).toString())
            .time(new SimpleDateFormat("yyyy年MM月dd日HH:mm:ss").parse(getCellValue(row.getCell(2)).toString())).build();
      users.add(user);
    }


    for(User user : users){
      System.out.println(user);
    }


    //把封装后的List<User>传给dao层,通过dao层来存入数据  ---> 此处演示省略...
  }


    //@GetMapping("/export")
  @RequestMapping(value = "/export",method = RequestMethod.POST)
  public String exportExcel(HttpServletResponse response, @RequestBody Map<String,Object> map) throws IOException {
      String username = map.get("username").toString();
      String password = map.get("password").toString();
      System.out.println(username);
      //1.创建Excel workbook
    //XSSFWorkbook  ---> 2007+
    //HSSFWorkbook ---> 2003
    XSSFWorkbook workbook = new XSSFWorkbook();
    //2.创建工作簿sheet
    Sheet sheet = workbook.createSheet("用户表");
    //3.创建标题头数组
    String [] arr = {"序号","姓名","入学时间"};
    //声明原子性对象
    AtomicInteger headAi = new AtomicInteger();
    //获得工作簿的第一行,用于存入标题头
    Row headRow = sheet.createRow(0);
    //第一行的标题头赋值
    for(String name : arr){
      Cell cell = headRow.createCell(headAi.getAndIncrement());
      cell.setCellValue(name);
    }

    AtomicInteger bodyAi = new AtomicInteger(1);

    Cell cell = null;

    List<User> users = this.createData();
    //4.存入数据
    for (User user : users) {
      Row bodyRow = sheet.createRow(bodyAi.getAndIncrement());
      cell = bodyRow.createCell(0);
      cell.setCellValue(user.getId());

      cell = bodyRow.createCell(1);
      cell.setCellValue(user.getName());

      cell = bodyRow.createCell(2);
      cell.setCellValue(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(user.getTime()));
    }

    String fileName = URLEncoder.encode("用户8月份数据.xlsx", "UTF-8");
    //5.下载
    response.setContentType("application/octet-stream");
    response.setHeader("content-disposition","attachment;filename=" + new String(fileName.getBytes("ISO-8859-1")));
    response.setHeader("filename",fileName);
    workbook.write(response.getOutputStream());

    System.out.println("导出完成");
    return username+password;

  }


  public List<User> createData(){
    List<User> users = new ArrayList<>();
    for(int i = 1; i <= 100; i ++){
      User user = User.builder().id(i + "").name("张三" + i).time(new Date()).build();
      users.add(user);
    }
    return users;
  }


  public static Object getCellValue(Cell cell){
    //1 获得单元格的属性类型
    CellType cellType = cell.getCellType();

    //2.根据单元格数据类型来处理数据
    Object value = null;

    switch (cellType){
      case STRING:

        value = cell.getStringCellValue();
        break;
      case BOOLEAN:
        value = cell.getBooleanCellValue();
        break;
      case NUMERIC:
        if(DateUtil.isCellDateFormatted(cell)){
          //日期格式
          value = cell.getDateCellValue();
        }else{
          //数字
          value = cell.getNumericCellValue();
        }
        break;
      case FORMULA: //公式
        value = cell.getCellFormula();
        break;
    }
    return value;
  }


  @GetMapping("/util/export")
  public void utilForExport(HttpServletResponse response) throws IOException, IllegalAccessException {
    ExcelExportUtil.export(response,this.createData(),User.class,"haha.xlsx","user.xlsx");
  }


  @GetMapping("/util/import")
  public void utilForImport(MultipartFile file) throws IOException, InstantiationException, IllegalAccessException, java.text.ParseException {
    List<Object> objects = new ExcelImportUtil<>(User.class).readExcel(User.class, file.getInputStream());
  }


  public List<EmployeeReportResult> createData1() {
    List<EmployeeReportResult> employeeReportResults = new ArrayList<>();
    for (int i = 0; i < 100; i++) {
      EmployeeReportResult employeeReportResult = new EmployeeReportResult();
      employeeReportResult.setAge("17");
      employeeReportResult.setArchivingOrganization("1");
      employeeReportResult.setAreThereAnyMajorMedicalHistories("1");
      employeeReportResult.setBankCardNumber("1");
      employeeReportResult.setBirthday("1");
      employeeReportResult.setBloodType("1");
      employeeReportResult.setCertificateOfAcademicDegree("1");
      employeeReportResult.setCompanyId("1");
      employeeReportResult.setConstellation("1");
      employeeReportResult.setContactTheMobilePhone("1");
      employeeReportResult.setDateOfBirth("1");
      employeeReportResult.setDateOfResidencePermit("1");
      employeeReportResult.setDepartmentName("1");
      employeeReportResult.setDoChildrenHaveCommercialInsurance("1");
      employeeReportResult.setDomicile("1");
      employeeReportResult.setEmergencyContact("1");
      employeeReportResult.setEducationalType("1");
      employeeReportResult.setEmergencyContactNumber("1");
      employeeReportResult.setEnglishName("1");
      employeeReportResult.setEnrolmentTime("1");
      employeeReportResult.setGraduateSchool("1");
      employeeReportResult.setGraduationCertificate("1");
      employeeReportResult.setGraduationTime("1");
      employeeReportResult.setHomeCompany("1");
      employeeReportResult.setIdCardPhotoBack("1");
      employeeReportResult.setIdCardPhotoPositive("1");
      employeeReportResult.setIdNumber("1");
      employeeReportResult.setIsThereAnyCompetitionRestriction("1");
      employeeReportResult.setIsThereAnyViolationOfLawOrDiscipline("1");
      employeeReportResult.setMajor("1");
      employeeReportResult.setMaritalStatus("1");
      employeeReportResult.setMobile("1");
      employeeReportResult.setMobile("1");
      employeeReportResult.setNation("1");
      employeeReportResult.setNationalArea("1");
      employeeReportResult.setNativePlace("1");
      employeeReportResult.setOpeningBank("1");
      employeeReportResult.setPassportNo("1");
      employeeReportResult.setPersonalMailbox("1");
      employeeReportResult.setPlaceOfResidence("1");
      employeeReportResult.setPoliticalOutlook("1");
      employeeReportResult.setPostalAddress("1");
      employeeReportResult.setProofOfDepartureOfFormerCompany("1");
      employeeReportResult.setProvidentFundAccount("1");
      employeeReportResult.setQq("1");
      employeeReportResult.setReasonsForLeaving("1");
      employeeReportResult.setRemarks("1");
      employeeReportResult.setResidenceCardCity("1");
      employeeReportResult.setReasonsForLeaving("1");
      employeeReportResult.setResidencePermitDeadline("1");
      employeeReportResult.setResignationTime("1");
      employeeReportResult.setResume("1");
      employeeReportResult.setSex("1");
      employeeReportResult.setSocialSecurityComputerNumber("1");
      employeeReportResult.setStaffPhoto("1");
      employeeReportResult.setStateOfChildren("1");
      employeeReportResult.setTheHighestDegreeOfEducation("1");
      employeeReportResult.setTimeOfEntry("1");
      employeeReportResult.setTimeToJoinTheParty("1");
      employeeReportResult.setTitle("1");
      employeeReportResult.setTypeOfTurnover("1");
      employeeReportResult.setUserId("1");
      employeeReportResult.setUsername("1");
      employeeReportResult.setWechat("1");
      employeeReportResult.setZodiac("1");
      employeeReportResults.add(employeeReportResult);
    }
    return employeeReportResults;
  }


  @GetMapping("/template")
  public void templateExport(HttpServletResponse response) throws IOException {

    Resource resource = new ClassPathResource("header-demo.xlsx");
    FileInputStream in = new FileInputStream(resource.getFile());


    //创建Excel工作对象
    XSSFWorkbook wb = new XSSFWorkbook(in);

    //读取工作表
    Sheet sheet = wb.getSheetAt(0);

    //抽取公共样式
    Row styleRow = sheet.getRow(2);

    CellStyle [] styles = new CellStyle[styleRow.getLastCellNum()];

    for(int i = 0; i < styleRow.getLastCellNum(); i ++){
      styles[i] = styleRow.getCell(i).getCellStyle();
    }

    //准备存入Excel表中的数据集合
    List<EmployeeReportResult> list = this.createData1();

    AtomicInteger dataAi = new AtomicInteger(2);

    Cell cell = null;

    for (EmployeeReportResult report : list) {

      Row dataRow = sheet.createRow(dataAi.getAndIncrement());

      //编号
      cell = dataRow.createCell(0);
      cell.setCellValue(report.getUserId());
      cell.setCellStyle(styles[0]);

      //姓名
      cell = dataRow.createCell(1);
      cell.setCellValue(report.getUsername());
      cell.setCellStyle(styles[1]);
      //手机
      cell = dataRow.createCell(2);
      cell.setCellValue(report.getMobile());
      cell.setCellStyle(styles[2]);
      //最高学历
      cell = dataRow.createCell(3);
      cell.setCellValue(report.getTheHighestDegreeOfEducation());
      cell.setCellStyle(styles[3]);

      //国家地区
      cell = dataRow.createCell(4);
      cell.setCellValue(report.getNationalArea());
      cell.setCellStyle(styles[4]);


      //护照
      cell = dataRow.createCell(5);
      cell.setCellValue(report.getPassportNo());
      cell.setCellStyle(styles[5]);

      //籍贯
      cell = dataRow.createCell(6);
      cell.setCellValue(report.getNativePlace());
      cell.setCellStyle(styles[6]);
      //生日
      cell = dataRow.createCell(7);
      cell.setCellValue(report.getBirthday());
      cell.setCellStyle(styles[7]);

      //属相
      cell = dataRow.createCell(8);
      cell.setCellValue(report.getZodiac());
      cell.setCellStyle(styles[8]);

      //入职时间
      cell = dataRow.createCell(9);
      cell.setCellValue(report.getTimeOfEntry());
      cell.setCellStyle(styles[9]);


      //离职类型
      cell = dataRow.createCell(10);
      cell.setCellValue(report.getTypeOfTurnover());
      cell.setCellStyle(styles[10]);
      //离职原因
      cell = dataRow.createCell(11);
      cell.setCellValue(report.getReasonsForLeaving());
      cell.setCellStyle(styles[11]);
      //离职时间
      cell = dataRow.createCell(12);
      cell.setCellValue(report.getResignationTime());
      cell.setCellStyle(styles[12]);

    }

    String filename = URLEncoder.encode("人缘信息.xlsx", "UTF-8");
    response.setContentType("application/octet-stream");
    response.setHeader("content-disposition","attachment;filename=" + new String(filename.getBytes("ISO-8859-1")));
    response.setHeader("filename",filename);
    wb.write(response.getOutputStream());


  }


  @GetMapping("/million/template")
  public void templateMillionExport(HttpServletResponse response) throws IOException {

    //2.创建工作簿
    SXSSFWorkbook workbook = new SXSSFWorkbook();
    //3.构造sheet
    String[] titles = {"编号", "姓名", "手机", "最高学历", "国家地区", "护照号", "籍贯",
        "生日", "属相", "入职时间", "离职类型", "离职原因", "离职时间"};
    Sheet sheet = workbook.createSheet();
    Row row = sheet.createRow(0);
    AtomicInteger headersAi = new AtomicInteger();
    for (String title : titles) {
      Cell cell = row.createCell(headersAi.getAndIncrement());
      cell.setCellValue(title);
    }
    AtomicInteger datasAi = new AtomicInteger(1);
    Cell cell = null;
    List<EmployeeReportResult> list = this.createData1();
    for (int i = 0; i < 10000; i++) {
      for (EmployeeReportResult report : list) {
        Row dataRow = sheet.createRow(datasAi.getAndIncrement());
        //编号
        cell = dataRow.createCell(0);
        cell.setCellValue(report.getUserId());
        //姓名
        cell = dataRow.createCell(1);
        cell.setCellValue(report.getUsername());
        //手机
        cell = dataRow.createCell(2);
        cell.setCellValue(report.getMobile());
        //最高学历
        cell = dataRow.createCell(3);
        cell.setCellValue(report.getTheHighestDegreeOfEducation());
        //国家地区
        cell = dataRow.createCell(4);
        cell.setCellValue(report.getNationalArea());
        //护照号
        cell = dataRow.createCell(5);
        cell.setCellValue(report.getPassportNo());
        //籍贯
        cell = dataRow.createCell(6);
        cell.setCellValue(report.getNativePlace());
        //生日
        cell = dataRow.createCell(7);
        cell.setCellValue(report.getBirthday());
        //属相
        cell = dataRow.createCell(8);
        cell.setCellValue(report.getZodiac());
        //入职时间
        cell = dataRow.createCell(9);
        cell.setCellValue(report.getTimeOfEntry());
        //离职类型
        cell = dataRow.createCell(10);
        cell.setCellValue(report.getTypeOfTurnover());
        //离职原因
        cell = dataRow.createCell(11);
        cell.setCellValue(report.getReasonsForLeaving());
        //离职时间
        cell = dataRow.createCell(12);
        cell.setCellValue(report.getResignationTime());
      }
    }
    String fileName = URLEncoder.encode("人员百万信息.xlsx", "UTF-8");
    response.setContentType("application/octet-stream");
    response.setHeader("content-disposition", "attachment;filename=" + new
        String(fileName.getBytes("ISO8859-1")));
    response.setHeader("filename", fileName);
    workbook.write(response.getOutputStream());

  }


}
