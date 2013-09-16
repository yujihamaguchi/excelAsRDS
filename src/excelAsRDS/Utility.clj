(ns ^{:author "Yuji Hamaguchi"}
  excelAsRDS.Utility
  (:use
    [clojure.data.json :as json :only [read-json write-str]]
    clojure.set)
  (:gen-class
    :name excelAsRDS.Utility
    :methods [
      #^{:static true} [isEqualJSONStrAsSet [String String] Boolean]
      #^{:static true} [differenceJSONStrAsSet [String String] String]
    ]))

(defn isEqualJSONStrAsSet
  "Returns true if arguments are equal to each other as set, false otherwise."
  [set1-json set2-json]
  (let [
    set1 (set (json/read-json set1-json))
    set2 (set (json/read-json set2-json))]
    (= set1 set2)))

(defn -isEqualJSONStrAsSet [set1-json set2-json]
  (isEqualJSONStrAsSet set1-json set2-json))

(defn differenceJSONStrAsSe
  "Returns different arguments. (set1 - set2)"
  [set1-json set2-json]
  (let [
    set1 (set (json/read-json set1-json))
    set2 (set (json/read-json set2-json))]
    (json/write-str (difference set1 set2))))

(defn -differenceJSONStrAsSet [set1-json set2-json]
  (differenceJSONStrAsSet set1-json set2-json))
