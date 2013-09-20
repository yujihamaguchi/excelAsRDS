# excelAsRDS

Manipulate excel spreadsheet data as relational data source(SELECT, INSERT, UPDATE).  
Based on [Apache POI](http://poi.apache.org), written in [Clojure](http://clojure.org).

## Building

    lein deps
    lein uberjar

## Example

* BeanShell

```
import excelAsRDS.*;

String outString = "";
String errMessage = "";
String errClassName = "";

try {
  // Select
  outString = Dml.selectSS(
    "./resources/test01.json" // define schema file in json
    ,"./resources/test01.xls" // excel file as datasource
    ,"{ \"attributes\" : [\"id\", \"pwd\"] \"whereClause\" { \"id\" : \"x\" }}" // query as json string
  );
  // Update
  Dml.updateSS(
    "./resources/test08.json" // define schema file in json
    ,"./resources/test12.xls" // excel file as datasource
    ,"{ \"attributes\" : [\"id\", \"pwd\"] \"whereClause\" { \"id\" : \"x\" }}" // query as json string
  );
  // Insert
  outString = Dml.selectSS(
    "./resources/test01.json" // define schema file in json
    ,"./resources/test01.xls" // excel file as datasource
    ,"[ { \"pwd\" : \"p11\", \"whereClause\" : { \"id\" : \"x\" } }
      , { \"pwd\" : \"p22\", \"whereClause\" : { \"id\" : \"y\" } } ]" // query as json string
  );
} catch (Exception e) {
  errMessage = e.getMessage();
  errClassName = e.getClass().getName();
}
```

* SCHEMA DIFINITION(json)

```
{
  "sheetIndex" : 0
  "columnIndex" : {
    "id" : 0
    "pwd" : 1
  }
  "startRowIndex" : 3
  "endRowIndex" : 5
  "required" : ["id"]
}
```

## License

Copyright © 2013 Yuji Hamaguchi

Released under the MIT license

https://github.com/yujihamaguchi/excelAsRDS/blob/master/license.txt
