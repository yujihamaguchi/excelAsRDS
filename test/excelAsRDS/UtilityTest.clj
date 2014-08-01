(ns excelAsRDS.UtilityTest
  (:require
    [clojure.test :refer :all]
    [excelAsRDS.Utility :refer :all]))

(deftest ut-if-lets
  (testing "if-lets macro (normal cases)"
    (is (= 0 (if-lets [x 0] x)))
    (is (= 0 (if-lets [x 0] x 1)))
    (is (= 1 (if-lets [x nil] x 1)))
    (is (= 0 (if-lets [x 0 y x] y)))
    (is (= 0 (if-lets [x 0 y x] y 1)))
    (is (= 1 (if-lets [x nil y x] y 1)))
    (is (= 0 (if-lets [x 0 y x z y] z)))
    (is (= 0 (if-lets [x 0 y x z y] z 1)))
    (is (= 1 (if-lets [x nil y x z y] y 1)))
    (is (= true (if-lets [x true] true false)))
    (is (= false (if-lets [x false] true false)))
    (is (= true (if-lets [x true y true] true false)))
    (is (= false (if-lets [x false y true] true false)))
    (is (= false (if-lets [x true y false] true false)))
    (is (= true (if-lets [x true y true z true] true false)))
    (is (= false (if-lets [x false y true z true] true false)))
    (is (= false (if-lets [x true y false z true] true false)))
    (is (= false (if-lets [x true y true z false] true false)))
  )
)

(deftest ut-if-lets-ab
  (testing "if-lets macro (abnormal cases)"
    (is (= (try (if-lets [] true false) (catch Exception e (.getMessage e)))
        "if-lets requires 2 or multiple of 2 forms in binding vector in user:1"))
    (is (= (try (if-lets [x] true false) (catch Exception e (.getMessage e)))
        "if-lets requires 2 or multiple of 2 forms in binding vector in user:1"))
    (is (= (try (if-lets [x true y] true false) (catch Exception e (.getMessage e)))
        "if-lets requires 2 or multiple of 2 forms in binding vector in user:1"))
  )
)

(deftest ut-isEqualJSONStrAsSet
  (testing "isEqualJSONStrAsSet(正常系)"
    (testing "0要素(オブジェクト)が等しい"
      (is
        (isEqualJSONStrAsSet
          "[]"
          "[]"
        )
      )
    )
    (testing "1要素(リテラル)が等しい"
      (is
        (isEqualJSONStrAsSet
          "[1]"
          "[1]"
        )
      )
    )
    (testing "1要素(リテラル)が等しくない1"
      (is
        (not (isEqualJSONStrAsSet
          "[1]"
          "[2]"
        ))
      )
    )
    (testing "1要素(リテラル)が等しくない2"
      (is
        (not (isEqualJSONStrAsSet
          "[1]"
          "[]"
        ))
      )
    )
    (testing "2要素(リテラル)が等しい"
      (is
        (isEqualJSONStrAsSet
          "[1 2]"
          "[1 2]"
        )
      )
    )
    (testing "2要素(リテラル)が等しくない1"
      (is
        (not (isEqualJSONStrAsSet
          "[1 2]"
          "[1 3]"
        ))
      )
    )
    (testing "2要素(リテラル)が等しくない2"
      (is
        (not (isEqualJSONStrAsSet
          "[2 2]"
          "[1 2]"
        ))
      )
    )
    (testing "1要素(オブジェクト)が等しい1"
      (is
        (isEqualJSONStrAsSet
          "[{\"id\" : \"x\"}]"
          "[{\"id\" : \"x\"}]"
        )
      )
    )
    (testing "1要素(オブジェクト)が等しい2"
      (is
        (isEqualJSONStrAsSet
          "[{\"id\" : \"x\"}]"
          "[{\"id\" : \"x\"}, {\"id\" : \"x\"}]"
        )
      )
    )
    (testing "1要素(オブジェクト)が等しくない1"
      (is
        (not (isEqualJSONStrAsSet
          "[{\"id\" : \"x\"}]"
          "[{\"id\" : \"y\"}]"
        ))
      )
    )
    (testing "1要素(オブジェクト)が等しくない2"
      (is
        (not (isEqualJSONStrAsSet
          "[{\"pwd\" : \"x\"}]"
          "[{\"id\" : \"x\"}]"
        ))
      )
    )
    (testing "2要素(オブジェクト)が等しい1"
      (is
        (isEqualJSONStrAsSet
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しい2"
      (is
        (isEqualJSONStrAsSet
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          "[{\"pwd\" : \"px\"}, {\"id\" : \"x\"}]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しくない1"
      (is
        (not (isEqualJSONStrAsSet
          "[{\"id\" : \"y\"}, {\"pwd\" : \"px\"}]"
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
        ))
      )
    )
    (testing "2要素(オブジェクト)が等しくない2"
      (is
        (not (isEqualJSONStrAsSet
          "[{\"identity\" : \"x\"}, {\"pwd\" : \"px\"}]"
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
        ))
      )
    )
    (testing "2要素(オブジェクト)が等しくない3"
      (is
        (not (isEqualJSONStrAsSet
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          "[{\"id\" : \"x\"}, {\"pwd\" : \"py\"}]"
        ))
      )
    )
    (testing "2要素(オブジェクト)が等しくない4"
      (is
        (not (isEqualJSONStrAsSet
          "[{\"host\" : \"hx\"}, {\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
        ))
      )
    )
    (testing "2要素(オブジェクト)が等しくない5"
      (is
        (not (isEqualJSONStrAsSet
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}, {\"host\" : \"hx\"}]"
        ))
      )
    )
    (testing "0要素(オブジェクト)が等しい"
      (is
        (isEqualJSONStrAsSet
          "[]"
          "[]"
        )
      )
    )
  )
)

(deftest ut-differenceJSONStrAsSet
  (testing "differenceJSONStrAsSet(正常系)"
    (testing "0要素(オブジェクト)が等しい"
      (is
        (=
          (differenceJSONStrAsSet
            "[]"
            "[]"
          )
          "[]"
        )
      )
    )
    (testing "1要素(リテラル)が等しい"
      (is
        (=
          (differenceJSONStrAsSet
            "[1]"
            "[1]"
          )
        )
        "[]"
      )
    )
    (testing "1要素(リテラル)が等しくない1"
      (is
        (=
          (differenceJSONStrAsSet
            "[1]"
            "[2]"
          )
          "[1]"
        )
      )
    )
    (testing "1要素(リテラル)が等しくない2"
      (is
        (=
          (differenceJSONStrAsSet
            "[1]"
            "[]"
          )
          "[1]"
        )
      )
    )
    (testing "2要素(リテラル)が等しい"
      (is
        (=
          (differenceJSONStrAsSet
            "[1 2]"
            "[1 2]"
          )
          "[]"
        )
      )
    )
    (testing "2要素(リテラル)が等しくない1"
      (is
        (=
          (differenceJSONStrAsSet
            "[1 2]"
            "[1 3]"
          )
          "[2]"
        )
      )
    )
    (testing "2要素(リテラル)が等しくない2"
      (is
        (=
          (differenceJSONStrAsSet
            "[2 2]"
            "[1 2]"
          )
          "[]"
        )
      )
    )
    (testing "1要素(オブジェクト)が等しい1"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"id\" : \"x\"}]"
            "[{\"id\" : \"x\"}]"
          )
          "[]"
        )
      )
    )
    (testing "1要素(オブジェクト)が等しい2"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"id\" : \"x\"}]"
            "[{\"id\" : \"x\"}, {\"id\" : \"x\"}]"
          )
          "[]"
        )
      )
    )
    (testing "1要素(オブジェクト)が等しくない1"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"id\" : \"x\"}]"
            "[{\"id\" : \"y\"}]"
          )
          "[{\"id\":\"x\"}]"
        )
      )
    )
    (testing "1要素(オブジェクト)が等しくない2"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"pwd\" : \"x\"}]"
            "[{\"id\" : \"x\"}]"
          )
          "[{\"pwd\":\"x\"}]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しい1"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          )
          "[]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しい2"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
            "[{\"pwd\" : \"px\"}, {\"id\" : \"x\"}]"
          )
          "[]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しくない1"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"id\" : \"y\"}, {\"pwd\" : \"px\"}]"
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          )
          "[{\"id\":\"y\"}]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しくない2"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"identity\" : \"x\"}, {\"pwd\" : \"px\"}]"
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          )
          "[{\"identity\":\"x\"}]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しくない3"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
            "[{\"id\" : \"x\"}, {\"pwd\" : \"py\"}]"
          )
          "[{\"pwd\":\"px\"}]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しくない4"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"host\" : \"hx\"}, {\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
          )
          "[{\"host\":\"hx\"}]"
        )
      )
    )
    (testing "2要素(オブジェクト)が等しくない5"
      (is
        (=
          (differenceJSONStrAsSet
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}]"
            "[{\"id\" : \"x\"}, {\"pwd\" : \"px\"}, {\"host\" : \"hx\"}]"
          )
          "[]"
        )
      )
    )
    (testing "0要素(オブジェクト)が等しい"
      (is
        (=
          (differenceJSONStrAsSet
            "[]"
            "[]"
          )
          "[]"
        )
      )
    )
  )
)