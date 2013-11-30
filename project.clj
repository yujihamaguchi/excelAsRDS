(defproject excelAsRDS "0.0.1"
  :dependencies [
    [org.clojure/clojure "1.5.1"]
    [org.clojure/data.json "0.2.2"]
    [org.apache.poi/poi "3.9"]
    [org.apache.poi/poi-ooxml "3.9"]
    [org.clojure/java.jdbc "0.3.0-alpha4"]
    [com.microsoft/sqljdbc4 "3.0"]
  ]
  :aot [
    excelAsRDS.Dml
    excelAsRDS.Utility
  ]
)