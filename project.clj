(defproject excelAsRDS "0.0.1"
  :dependencies [
    [org.clojure/clojure "1.6.0"]
    [org.clojure/data.json "0.2.2"]
    [org.apache.poi/poi "3.9"]
    [org.apache.poi/poi-ooxml "3.9"]
  ]
  :plugins [[codox "0.6.6"]]
  :aot [
    excelAsRDS.Dml
    excelAsRDS.Utility
  ]
  :jvm-opts ["-Xmx768m"] 
)