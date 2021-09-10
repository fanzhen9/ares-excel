# Ares-Spring-boot-excel

### 说明

本软件采用springboot 2.3.7.RELEASE 开发，继承POI 5.0.0

### 使用方法

* springboot 工程集成

  首先在pom.xml文件中添加依赖

  ```xml
  <dependency>
       <groupId>org.springframework.boot</groupId>
       <artifactId>spring-boot-starter-web-ares-excel</artifactId>
       <version>0.0.1-SNAPSHOT</version>
  </dependency>
  ```

* 配置文件

  ```properties
  # 生成文件路径
  ares.excel.filePath=d://excel/
  # id生成类，可以不配置，即生成出来的文件名称
  ares.excel.idGenerate=com.example.APP
  ```

* 注解使用

  ```java
  @ExcelSheet(name = "学生信息") //sheet 名称
  public class Student {
  	
      //列序号  表头
      @Excel(index = 0,name = "班级") 
      private String className;
  
      @Excel(index = 1,name = "成绩")
      private Integer score;
  
      @Excel(index = 2,name="盆友")
      private String friend;
  	//列序号  表头 是否是图片,如果是图片则下载 字段类型 读取时有效 下载图片自定义方法
      @Excel(index = 3,name = "图片",isUrl = true,fieldClassType = String.class,downLoadImageService = MyDownLoadService.class)
      private String url;
  }
  ```

  方法使用

  ```java
  List<Data> list = new ArrayList<>();
  List<Teacher> teachers = new ArrayList<>();
  teachers.add(new Teacher("王老师",28,"足球"));
  teachers.add(new Teacher("张老师",25,"乒乓球"));
  teachers.add(new Teacher("李老师",28,"羽毛球"));
  Data data = new Data();
  data.setList(teachers);
  data.setClazz(Teacher.class);
  
  List<Student> students = new ArrayList<>();
  students.add(new Student("三年二班",60,"小红","http://localhost:8080/aaa.jpg"));
  students.add(new Student("三年一班",50,"小张","http://localhost:8080/aaa.jpg"));
  students.add(new Student("四年二班",70,"王强","http://localhost:8080/aaa.jpg"));
  Data data1 = new Data();
  data1.setList(students);
  data1.setClazz(Student.class);
  
  
  list.add(data);
  list.add(data1);
  excelService.writeExcel(list);
  ```

  也可以使用模板方式,提前准备好需要的模板

  ![image-20210909152021206](C:\Users\fox\AppData\Roaming\Typora\typora-user-images\image-20210909152021206.png)

  只要和类字段保持一致，就可以导出 格式是{.xxx}

  ```java
  List<Data> list = new ArrayList<>();
  List<Teacher> teachers = new ArrayList<>();
  teachers.add(new Teacher("王老师",28,"足球"));
  teachers.add(new Teacher("张老师",25,"乒乓球"));
  teachers.add(new Teacher("李老师",28,"羽毛球"));
  Data data = new Data();
  data.setList(teachers);
  data.setClazz(Teacher.class);
  
  List<Student> students = new ArrayList<>();
  students.add(new Student("三年二班",60,"小红","http://localhost:8080/aaa.jpg"));
  students.add(new Student("三年一班",50,"小张","http://localhost:8080/aaa.jpg"));
  students.add(new Student("四年二班",70,"王强","http://localhost:8080/aaa.jpg"));
  Data data1 = new Data();
  data1.setList(students);
  data1.setClazz(Student.class);
  
  
  list.add(data);
  list.add(data1);
  FileInputStream fileInputStream = new FileInputStream("d://a.xlsx");
  excelService.writeExcel(fileInputStream,list,"123456");
  ```



* 使用自定义下载

  需要继承AbstractDownLoadImageService 并在注解中配置自定义下载

  ```java
  public class MyDownLoadService extends AbstractDownLoadImageService {
      public MyDownLoadService(String path, String fileName, String url, CountDownLatch countDownLatch, RestTemplate restTemplate) {
          super(path, fileName, url, countDownLatch, restTemplate);
      }
  
      @Override
      public void downLoadImage() {
          System.out.println(path);
          System.out.println(fileName);
          System.out.println(url);
          System.out.println(countDownLatch);
          countDownLatch.countDown();
          System.out.println(restTemplate);
      }
  }
  ```

* 自定义文件名生成策略

  实现IdGenerate接口并返回一个id

  并且在配置文件中配置这个类

  ```java
  public class Mygen implements IdGenerate {
      @Override
      public String getName() {
          return null;
      }
  }
  ```

  

