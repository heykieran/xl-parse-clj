# Convert an Excel Workbook to Clojure Code

## Introduction

This is a very rough proof of concept of converting a non-trivial, but not overly complicated, Excel
workbook to Clojure code. It is just to scratch an itch, and I've made no effort yet to clean it up
and turn it into a library. I'm making it available as is. If it's useful to you, that's great, but 
please realize that there's a long way to go before it is any more than merely marginally useful.

Over time, as that time becomes available to me, I'll expand it and refine it, but I make no 
committments as to when that will be. Also, I've given little consideration to performance or catching
exceptions.

**Use anything here at your own risk.**

At the moment, the only external dependencies, in addition to Clojure itself, are [docjure](https://github.com/mjul/docjure), [ubergraph](https://github.com/Engelberg/ubergraph), and 
[numeric-tower](https://github.com/clojure/math.numeric-tower), all excellent libraries.

**It is a work in progress, so you'll find a lot of `comment` forms scattered about the code. These may or
may not work at any particular time, but I find them useful**

## Convert Excel formulas to AST

Much of the effort of extracting an AST from an individual Excel formula is based on some excellent prior work done by E. W. Bachtal and can be found [here](https://ewbi.blogs.com/develops/). 

The algorithm was minimally "functionalized" and converted to Clojure, but I acknowledge
a debt of gratitude to the original author.

I considered ANTLR and other similar approaches, but rejected them because of the difficulty of dealing with
all the quirks and edge cases of Excel formulas.

Most of the action occurs on the `parse/parse-to-tokens` function, which when given an Excel
formula as a string will return the AST e.g.

```clojure
(-> "=1+2"
    (parse/parse-to-tokens))
```

will return

```clojure
[{:value "1", :type :Operand, :sub-type :Number}
 {:value "+", :type :OperatorInfix, :sub-type :Math}
 {:value "2", :type :Operand, :sub-type :Number}]
```

and

```clojure
(-> "=max(1,2)*$A$4"
    (parse/parse-to-tokens))
```

returns

```clojure
[{:value "max", :type :Function, :sub-type :Start}
 {:value "1", :type :Operand, :sub-type :Number}
 {:value ",", :type :Argument, :sub-type nil}
 {:value "2", :type :Operand, :sub-type :Number}
 {:value "", :type :Function, :sub-type :Stop}
 {:value "*", :type :OperatorInfix, :sub-type :Math}
 {:value "$A$4", :type :Operand, :sub-type :Range}]
```

This gives the AST for a single Excel formula, but to be more useful, it's helpful to provide
some form of structural nesting.

The `parse/nest-ast` function provides such a facility, for example,

```clojure
(-> "=max(1,2)*$A$4"
    (parse/parse-to-tokens)
    (parse/nest-ast))
```

returns the AST as a nested structure, below,

```clojure
{0
 {1
  {:sub-type :Start,
   :type :Function,
   :value "max",
   2 {:sub-type :Number, :type :Operand, :value "1"},
   3 {:sub-type nil, :type :Argument, :value ","},
   4 {:sub-type :Number, :type :Operand, :value "2"}},
  6 {:sub-type :Math, :type :OperatorInfix, :value "*"},
  7 {:sub-type :Range, :type :Operand, :value "$A$4"}}}
```

## Convert formula AST to equivalent Clojure code

Now that we have a suitable AST, how do we convert it to code? The code uses a variation of
the shunting-yard alogrithm where the precendences and associations are defined in the vector
`shunting/OPERATORS-PRECEDENCE` and where the functions themselves are defined in the `functions`
namespace or taken from `clojure.core`.

For example the `max` function is defined in the `shunting/OPERATORS-PRECEDENCE` vector as follows

```clojure
{:name :max :s "max" :f 'max :c :args :a :all :e [:Function :Start]}
```

which maps the excel formula symbol `"max"` to the clojure symbol `max`, which exists in the 
`clojure.core` namespace.

The `average` function is defined as follows

```clojure
{:name :average :s "average" :f `fns/average :c :args :a :all :e [:Function :Start]}
```

which maps the excel formula symbol `"average"` to the clojure `fns/average` symbol that is 
defined in the `functions` namespace and aliased to `fns`.

The order of the definitions in the `OPERATORS-PRECEDENCE` vector also determines the precedence
of the operators when being processed by the shunting-yard algorithm. Furthermore, the `:a` field
in the definition lets the parser know how many arguments to expect.

> As of this writing all Excel's mathematical and logical operators have been implemented, as well
  as the following functions `abs`, `sin`, `true`, `false`, `and`, `or`, `max`, `min`, `pi`, `now`, `date`, `days`, `datevalue`, `yearfrac`,
  `sum`, `average`, `count`, `counta` & `if`. Others will be added as necessary.
  
> A preliminary implementation of `vlookup` is also included. This will be extended.


As an example 

```clojure
(-> "=max(1,2)*$A$4"
    (parse/parse-to-tokens)
    (parse/nest-ast)
    (parse/wrap-ast)
    (ast/process-tree)
    (sh/parse-expression-tokens)
    (ast/unroll-for-code-form "Sheet1"))
```

yields

```clojure
(* (max 1.0 2.0) (eval-range "Sheet1!$A$4"))
```

### Testing the formulas

Included as part of this repository in the `resources` directory is an Excel Workbook **TEST1.xlsx**.

The sheet **Sheet1** of the workbook contains a number of formulas with static input arguments i.e. arguments that  do **not** refer to other cells or named ranges.

We can use this to verify that the clojure code being generated, when executed, returns the same
value as Excel.

The `core/run-tests` function does this. It inspects all the formulas in the second column of **Sheet1**,
converts them to clojure code, and evaluates that code. The function returns a vector of maps, one per formula,
with a key `:ok?` which will be set to true if the value returned by the clojure evaluation
is equal to the value calculated by Excel.

So, executing 

```clojure
(->> (run-tests)
       (filter #(false? (:ok? %))))
```

should return an empty sequence, if all the tests pass.

## Working with Excel Workbooks

Now, we're at a point where we can convert an Excel formula to a reasonable Clojure representation, but 
we still need to solve for two problems.

1. Referencing the Cell or Named Range values used in a formula (e.g. `$A$1` or `EMPLOYEES`), and
2. Determining the calculation order for the workbook so that dependencies are calculated before the cells
that depend on them.

With respect to item 1, as you've seen above, the formula

```"=max(1,2)*$A$4"```

is converted to to the Clojure form

```clojure
(* (max 1.0 2.0) (eval-range "Sheet1!$A$4"))
```

so, in order to proceed further, we'll need to both parse the workbook, and provide an implementation for `eval-range` that's aware of the values in the workbook.

We'll treat a workbook as a DAG, where the DAG's edges link cells to dependent cells, and also provide a way 
to resolve named ranges to the cells to which they refer.

The `graph/explain-workbook` provides base-level functionality to parse a workbook, returning a map with
`:named-ranges` and `:cells` keys which respectively provide information about the named ranges and the 
cells in the workbook. Each individual entry will also, in its `:references` value provide information
about what other cells this cell references.

For example, a cell `Sheet2!E6` might contain the formula `=SUM(B2:B4)-SUM(C2:C4)` which Excel evaluates
to the value 528. 

In that case, its matching entry in the `:cells` vector would be 

```clojure
{:format "General",
   :value 528.0,
   :type :numeric,
   :references
   [{:sheet "Sheet2", :label "B2", :type :general}
    {:sheet "Sheet2", :label "B3", :type :general}
    {:sheet "Sheet2", :label "B4", :type :general}
    {:sheet "Sheet2", :label "C2", :type :general}
    {:sheet "Sheet2", :label "C3", :type :general}
    {:sheet "Sheet2", :label "C4", :type :general}],
   :sheet "Sheet2",
   :column 5,
   :label "E6",
   :formula "SUM(B2:B4)-SUM(C2:C4)",
   :row 5}
```

You can see that the `:references` value contains the information about each cell on which the formula
is dependent.

In order to *"break out"* a sheet for a workbook you can run

```clojure
(graph/explain-workbook "TEST1.xlsx" "Sheet2")
```

Once we have this information we can begin the process of converting it to a DAG. 

First, we use `graph/get-cell-dependencies` to augment the map returned by `explain-workbook` with a `:dependencies` key where its value is a vector of 2-tuples where the first value in the tuple is a cell and the 
second value is the cells on which it depends.

We follow this with a call to `graph/add-graph` which uses the `:dependencies` key to construct the DAG. The
graph is added to the workbook map as a `:map` entry.

If you have graphviz installed you can inspect the DAG produced by the test workbook's second worksheet
as follows

```clojure
(-> "TEST1.xlsx"
    (explain-workbook "Sheet2")
    (get-cell-dependencies)
    (add-graph)
    :graph
    (uber/viz-graph))
```

#### Dependencies of demo Workbook

![Dependency Graph](/assets/DAG1.png "Dependencies")

It still remains to provide an implementation of the `eval-range` function that is returned in the Clojure expression for any formula cell that references one, or more, other cells.

The code provides a function `graph/expand-cell-range` which when given a string describing a cell, a range 
of cells, or a named range in a workbook will expand it to the individual cells referenced.

As an example, if we `def` a variable to contain the *worksheet map* as follows:

```clojure
(def WB-MAP
    (-> "TEST1.xlsx"
        (explain-workbook "Sheet2")
        (get-cell-dependencies)
        (add-graph)))
```

we can the use

```clojure
(expand-cell-range "Sheet2!B3:D3" (:named-ranges WB-MAP))
```

to return information about the range "Sheet2!B3:D3", which returns

```clojure
[{:sheet "Sheet2", :label "B3", :type :general}
 {:sheet "Sheet2", :label "C3", :type :general}
 {:sheet "Sheet2", :label "D3", :type :general}]
```

and for a named range

```clojure
(expand-cell-range "Sheet2!BONUS" (:named-ranges WB-MAP))
```

returns the cell, or cells, to which the named range refers

```clojure
[{:sheet "Sheet2", :label "B7", :type :general}]
```

Building on this is the actual `graph/eval-range` function

```clojure
(eval-range "Sheet2!C2:C4" WB-MAP)
```

which returns

```clojure
[12.0 24.0 36.0]
```

which is a vector of the values contained in the range. 

Notice that _even_ for ranges that 
describe a rectangular region (rather than a single row or a single column) `eval-range` 
returns a vector.

However, `eval-range` also attaches meta-data to the vector returned, so that the _shape_
of the range can be recovered and used by functions that expect tabular data. For example

```clojure
(eval-range "Sheet2!$L$4:$N$6" WB-MAP)
```

returns

```clojure
["L1" 0.1 0.0 "L2" 0.2 30.0 "L3" 0.3 35.0]
```

and

```clojure
(meta (eval-range "Sheet2!$L$4:$N$6" WB-MAP))
```

returns

```clojure
{:single? false, :column? false, :sheet-name "Sheet2", 
 :tl-name "L4", :tl-coord [3 11], :cols 3, :rows 3}
```

which describes how the vector can be converted to a table.

> The function `expand-cell-range` adds the meta-data that is recapitulated by
  `eval-range`

So, finally, the `graph` function will walk the DAG and recalculate, in the correct order, 
the entire workbook using the clojure code that was generated during initial processing.

So, calling the following to recalculate Sheet2 of the example workbook

```clojure
(recalc-workbook WB-MAP "Sheet2")
```

will return a vector of tuples, where each tuple is the results of recalculating each formula
cell and contains the cell reference, a boolean indicating whether the calculated value is equal
to the cached value calculated by Excel, the text of the formula, the cached Excel result, the 
clojure form representing the Excel formula, and the value calculated by evaluating the Clojure
code.

For example, using the demo workbook, we get

```clojure
[["Sheet2!C3" true "BONUS * B3" 24.0 (* (eval-range "Sheet2!BONUS") (eval-range "Sheet2!B3")) 24.0]
 ["Sheet2!E3" true "SUM(B3:D3)" 229.0 (functions/sum (eval-range "Sheet2!B3:D3")) 229.0]
 ["Sheet2!C2" true "BONUS * B2" 12.0 (* (eval-range "Sheet2!BONUS") (eval-range "Sheet2!B2")) 12.0]
 ["Sheet2!B9" true "COUNTA(EMPLOYEES)" 3.0 (functions/fn-counta (eval-range "Sheet2!EMPLOYEES")) 3.0]
 ["Sheet2!B5" true "SUM(B2:B4)" 600.0 (functions/sum (eval-range "Sheet2!B2:B4")) 600.0]
 ["Sheet2!C4" true "BONUS * B4" 36.0 (* (eval-range "Sheet2!BONUS") (eval-range "Sheet2!B4")) 36.0]
 ["Sheet2!E6"
  true
  "SUM(B2:B4)-SUM(C2:C4)"
  528.0
  (- (functions/sum (eval-range "Sheet2!B2:B4")) (functions/sum (eval-range "Sheet2!C2:C4")))
  528.0]
 ["Sheet2!C5" true "SUM(C2:C4)" 72.0 (functions/sum (eval-range "Sheet2!C2:C4")) 72.0]
 ["Sheet2!E4" true "SUM(B4:D4)" 340.0 (functions/sum (eval-range "Sheet2!B4:D4")) 340.0]
 ["Sheet2!D5" true "SUM(D2:D4)" 15.0 (functions/sum (eval-range "Sheet2!D2:D4")) 15.0]
 ["Sheet2!E2" true "SUM(B2:D2)" 118.0 (functions/sum (eval-range "Sheet2!B2:D2")) 118.0]
 ["Sheet2!E5" true "SUM(E2:E4)" 687.0 (functions/sum (eval-range "Sheet2!E2:E4")) 687.0]
 ["Sheet2!B11"
  true
  "IF(E5<ALLOWEDTOTAL,\"YES\",\"NO\")"
  "YES"
  (if (< (eval-range "Sheet2!E5") (eval-range "Sheet2!ALLOWEDTOTAL")) (str "YES") (str "NO"))
  "YES"]]
```

If we want to check that the Clojure results match the results returned by Excel, we can run

```clojure
  (keep (fn [[cell-label match? cell-formula cell-value cell-code calculated-value :as calc]]
          (when-not match?
            calc))
        (recalc-workbook WB-MAP "Sheet2"))
```

and expect to get an empty list `'()` returned.

## Future Work

Obviously, now that we can calculate a workbook, it would be nice to be able to update input cell
values and then recalculate those portions of the workbook that are effected, i.e. those cells
whose value are at some level dependent on the updated cell.

```clojure
(get-recalc-node-sequence "Sheet2!A2" WB-MAP)
```

will return a list of cells, in the correct order, that should be recalculated when the A2 cell on 
Sheet2 is updated

```clojure
'("Sheet2!A2" "Sheet2!B9")
```



