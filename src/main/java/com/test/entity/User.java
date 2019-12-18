package com.test.entity;

import com.test.utils.ExcelAttribute;
import lombok.*;

import java.util.Date;

@Data   //get-set方法
@NoArgsConstructor //空参构造
@AllArgsConstructor //满参构造
@ToString  //toString方法
@Builder //链式编程
public class User {

  @ExcelAttribute(sort = 0)
  private String id;

  @ExcelAttribute(sort = 1)
  private String name;

  @ExcelAttribute(sort = 2)
  private Date time;

  private String address;

  private Date createTime;

  private Date modifyTime;


}
