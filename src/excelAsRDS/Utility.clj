(ns ^{:doc "Operate excel sheet as relational data source using Apache POI."
      :author "Yuji Hamaguchi"}
  excelAsRDS.Utility
  (:gen-class
    :name excelAsRDS.Utility
    :methods [
      #^{:static true} [isEqualJSONStrAsSet [String String] Boolean]
      #^{:static true} [differenceJSONStrAsSet [String String] String]
    ]))

(use '[clojure.data.json :as json :only [read-json write-str]])
(use 'clojure.set)

(defn isEqualJSONStrAsSet [json-str1 json-str2]
  (let [
    set1 (set (json/read-json json-str1))
    set2 (set (json/read-json json-str2))]
    (= set1 set2)))

(defn -isEqualJSONStrAsSet [json-str1 json-str2]
  (isEqualJSONStrAsSet json-str1 json-str2))

(defn differenceJSONStrAsSet [json-str1 json-str2]
  (let [
    set1 (set (json/read-json json-str1))
    set2 (set (json/read-json json-str2))]
    (json/write-str (difference set1 set2))))

(defn -differenceJSONStrAsSet [json-str1 json-str2]
  (differenceJSONStrAsSet json-str1 json-str2))
