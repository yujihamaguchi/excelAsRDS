(ns ^{:author "Yuji Hamaguchi"}
  excelAsRDS.Utility
  (:require [clojure.data.json :as json :only [read-json write-str]]
            [clojure.java.io :as io :only [reader]]
            [clojure.set :refer :all])
  (:gen-class
    :name excelAsRDS.Utility
    :methods [#^{:static true} [isEqualJSONStrAsSet [String String] Boolean]
              #^{:static true} [differenceJSONStrAsSet [String String] String]]))

;;; macros
(defmacro if-lets
  ([bindings true-expr] `(if-lets ~bindings ~true-expr nil))
  ([bindings true-expr false-expr]
    (cond
      (or (not (seq bindings)) (not (zero? (rem (count bindings) 2))))
        `(throw (IllegalArgumentException. "if-lets requires 2 or multiple of 2 forms in binding vector in user:1"))
      (seq (drop 2 bindings))
        `(if-let ~(vec (take 2 bindings))
                 (if-lets ~(vec (drop 2 bindings))
                          ~true-expr
                          ~false-expr)
                 ~false-expr)
      :else
        `(if-let ~(vec bindings)
                 ~true-expr
                 ~false-expr))))

(defn isEqualJSONStrAsSet
  "Returns true if arguments are equal to each other as set, false otherwise."
  [set1-json set2-json]
  (let [set1 (set (json/read-json set1-json))
        set2 (set (json/read-json set2-json))]
    (= set1 set2)))

(defn -isEqualJSONStrAsSet [set1-json set2-json]
  (isEqualJSONStrAsSet set1-json set2-json))

(defn differenceJSONStrAsSet
  "Returns different arguments. (set1 - set2)"
  [set1-json set2-json]
  (let [set1 (set (json/read-json set1-json))
        set2 (set (json/read-json set2-json))]
    (json/write-str (difference set1 set2))))

(defn -differenceJSONStrAsSet [set1-json set2-json]
  (differenceJSONStrAsSet set1-json set2-json))
