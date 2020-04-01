# excel-helper
excel帮助类


本文代码ExcelReader部分模仿自cc.aicode.e2e（该库基于org.apache.poi）
https://aicode.cc/excel2entity
实现了Java POI对xls文件的读取功能的封装，实现了批量导入Excel中的数据时自动根据Excel中的数据行创建对应的Java POJO实体对象的功能。
该类库也实现了在创建实体对象时对字段类型进行校验，可以对Excel中的数据类型合法性进行校验，通过实现扩展接口，可以实现自定义校验规则以及自定义实体对象字段类型等更加复杂的校验规则和字段类型转换。
 
<dependency>
    <groupId>cc.aicode.java.e2e</groupId>
    <artifactId>ExcelToEntity</artifactId>
    <version>1.0.0.3</version>
</dependency>
 
 
优点：
核心，解析为Entity（采用反射机制）。
支持中文列名读取（使用自定义annotation进行列的解析）。
支持自定义字段类型（需要继承ExcelType，已经调用registerNewType方法进行注册），方便特殊字段value的处理（解析失败会给出null值）。
支持字段格式（正则表达式）验证。
验证必填字段（只检查，改列是否存在）。 
 
缺点：
暂无