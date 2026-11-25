Attribute VB_Name = "LispBuiltins"
Option Explicit
' ===========================================================================
'  LispBuiltins - primitive procedures: numbers, pairs and lists, strings,
'  characters, symbols, vectors, hash tables, records, control and output.
'  Every primitive is a LispObj (LK_PRIM) with Code = LP_*, I1 = minimum and
'  I2 = maximum number of arguments (-1 = any).
' ===========================================================================

' numbers 100..199
Private Const P_ADD As Long = 100
Private Const P_SUB As Long = 101
Private Const P_MUL As Long = 102
Private Const P_DIV As Long = 103
Private Const P_NUMEQ As Long = 104
Private Const P_LT As Long = 105
Private Const P_GT As Long = 106
Private Const P_LE As Long = 107
Private Const P_GE As Long = 108
Private Const P_QUOTIENT As Long = 109
Private Const P_REMAINDER As Long = 110
Private Const P_MODULO As Long = 111
Private Const P_ABS As Long = 112
Private Const P_MIN As Long = 113
Private Const P_MAX As Long = 114
Private Const P_GCD As Long = 115
Private Const P_LCM As Long = 116
Private Const P_EXPT As Long = 117
Private Const P_EXP As Long = 118
Private Const P_LOG As Long = 119
Private Const P_SIN As Long = 120
Private Const P_COS As Long = 121
Private Const P_TAN As Long = 122
Private Const P_ASIN As Long = 123
Private Const P_ACOS As Long = 124
Private Const P_ATAN As Long = 125
Private Const P_SQRT As Long = 126
Private Const P_SQUARE As Long = 127
Private Const P_NUMBERP As Long = 128
Private Const P_INTEGERP As Long = 129
Private Const P_RATIONALP As Long = 130
Private Const P_EXACTP As Long = 131
Private Const P_INEXACTP As Long = 132
Private Const P_EXACTINTP As Long = 133
Private Const P_ZEROP As Long = 134
Private Const P_POSITIVEP As Long = 135
Private Const P_NEGATIVEP As Long = 136
Private Const P_ODDP As Long = 137
Private Const P_EVENP As Long = 138
Private Const P_INEXACT As Long = 139
Private Const P_EXACT As Long = 140
Private Const P_FLOOR As Long = 141
Private Const P_CEILING As Long = 142
Private Const P_ROUND As Long = 143
Private Const P_TRUNCATE As Long = 144
Private Const P_NUM2STR As Long = 145
Private Const P_STR2NUM As Long = 146
Private Const P_RANDOM As Long = 147
Private Const P_ADD1 As Long = 148
Private Const P_SUB1 As Long = 149
Private Const P_NANP As Long = 150
Private Const P_EXACTSQRT As Long = 151
' lists 200..299
Private Const P_CONS As Long = 200
Private Const P_CAR As Long = 201
Private Const P_CDR As Long = 202
Private Const P_SETCAR As Long = 203
Private Const P_SETCDR As Long = 204
Private Const P_CXR As Long = 205
Private Const P_LIST As Long = 206
Private Const P_CONSSTAR As Long = 207
Private Const P_LENGTH As Long = 208
Private Const P_APPEND As Long = 209
Private Const P_REVERSE As Long = 210
Private Const P_LISTTAIL As Long = 211
Private Const P_LISTREF As Long = 212
Private Const P_LISTCOPY As Long = 213
Private Const P_LASTPAIR As Long = 214
Private Const P_LAST As Long = 215
Private Const P_MEMQ As Long = 216
Private Const P_MEMV As Long = 217
Private Const P_MEMBER As Long = 218
Private Const P_ASSQ As Long = 219
Private Const P_ASSV As Long = 220
Private Const P_ASSOC As Long = 221
Private Const P_MAP As Long = 222
Private Const P_FOREACH As Long = 223
Private Const P_FILTER As Long = 224
Private Const P_REMOVE As Long = 225
Private Const P_REDUCE As Long = 226
Private Const P_FOLDLEFT As Long = 227
Private Const P_FOLDRIGHT As Long = 228
Private Const P_FOLD As Long = 229
Private Const P_DELETE As Long = 230
Private Const P_LISTINDEX As Long = 231
Private Const P_IOTA As Long = 232
Public Const LP_APPLY As Long = 233
Private Const P_SORT As Long = 234
Private Const P_LISTSORT As Long = 235
Private Const P_ANY As Long = 236
Private Const P_EVERY As Long = 237
Private Const P_COUNT As Long = 238
Private Const P_DELDUP As Long = 239
Private Const P_APPENDMAP As Long = 240
Private Const P_FIND As Long = 241
Private Const P_LISTP As Long = 242
Private Const P_PAIRP As Long = 243
Private Const P_NULLP As Long = 244
Private Const P_MAKELIST As Long = 245
Private Const P_LISTSET As Long = 246
Private Const P_FILTERMAP As Long = 247
Private Const P_PARTITION As Long = 248
Private Const P_ASSOCSET As Long = 249
' strings, characters, symbols 300..399
Private Const P_SYMBOLP As Long = 300
Private Const P_STRINGP As Long = 301
Private Const P_CHARP As Long = 302
Private Const P_SYM2STR As Long = 303
Private Const P_STR2SYM As Long = 304
Private Const P_GENSYM As Long = 305
Private Const P_STRLEN As Long = 306
Private Const P_STRREF As Long = 307
Private Const P_SUBSTRING As Long = 308
Private Const P_STRAPPEND As Long = 309
Private Const P_STRCOPY As Long = 310
Private Const P_STREQ As Long = 311
Private Const P_STRLT As Long = 312
Private Const P_STRGT As Long = 313
Private Const P_STRLE As Long = 314
Private Const P_STRGE As Long = 315
Private Const P_STRCIEQ As Long = 316
Private Const P_STRUP As Long = 317
Private Const P_STRDOWN As Long = 318
Private Const P_STR2LIST As Long = 319
Private Const P_LIST2STR As Long = 320
Private Const P_STRING As Long = 321
Private Const P_MAKESTR As Long = 322
Private Const P_STRNULLP As Long = 323
Private Const P_STRINDEX As Long = 324
Private Const P_STRCONTAINS As Long = 325
Private Const P_STRSPLIT As Long = 326
Private Const P_STRJOIN As Long = 327
Private Const P_STRTRIM As Long = 328
Private Const P_STRTRIML As Long = 329
Private Const P_STRTRIMR As Long = 330
Private Const P_STRREV As Long = 331
Private Const P_STRPREFIXP As Long = 332
Private Const P_STRSUFFIXP As Long = 333
Private Const P_STRFOREACH As Long = 334
Private Const P_STRMAP As Long = 335
Private Const P_CHAR2INT As Long = 336
Private Const P_INT2CHAR As Long = 337
Private Const P_CHARUP As Long = 338
Private Const P_CHARDOWN As Long = 339
Private Const P_CHARALPHAP As Long = 340
Private Const P_CHARNUMP As Long = 341
Private Const P_CHARWSP As Long = 342
Private Const P_CHARUPPERP As Long = 343
Private Const P_CHARLOWERP As Long = 344
Private Const P_DIGITVAL As Long = 345
Private Const P_CHAREQ As Long = 346
Private Const P_CHARLT As Long = 347
Private Const P_CHARGT As Long = 348
Private Const P_CHARLE As Long = 349
Private Const P_CHARGE As Long = 350
Private Const P_STRPAD As Long = 351
Private Const P_STRPADR As Long = 352
Private Const P_STRCONTAINSP As Long = 353
Private Const P_STRREPLACE As Long = 354
' vectors and hash tables 400..499
Private Const P_VECTORP As Long = 400
Private Const P_MAKEVECTOR As Long = 401
Private Const P_VECTOR As Long = 402
Private Const P_VECREF As Long = 403
Private Const P_VECSET As Long = 404
Private Const P_VECLEN As Long = 405
Private Const P_VEC2LIST As Long = 406
Private Const P_LIST2VEC As Long = 407
Private Const P_VECFILL As Long = 408
Private Const P_VECMAP As Long = 409
Private Const P_VECFOREACH As Long = 410
Private Const P_VECCOPY As Long = 411
Private Const P_MAKEHASH As Long = 420
Private Const P_HASHP As Long = 421
Private Const P_HASHSET As Long = 422
Private Const P_HASHREF As Long = 423
Private Const P_HASHREFDEF As Long = 424
Private Const P_HASHGUILEREF As Long = 425
Private Const P_HASHDEL As Long = 426
Private Const P_HASHHAS As Long = 427
Private Const P_HASHCOUNT As Long = 428
Private Const P_HASHKEYS As Long = 429
Private Const P_HASHVALUES As Long = 430
Private Const P_HASH2ALIST As Long = 431
Private Const P_HASHUPDATE As Long = 432
Private Const P_HASHUPDATEDEF As Long = 433
Private Const P_HASHWALK As Long = 434
Private Const P_HASHCLEAR As Long = 435
Private Const P_HASHCOPY As Long = 436
Private Const P_HASHFOREACH As Long = 437
' control and output 500..599
Private Const P_PROCP As Long = 500
Private Const P_BOOLP As Long = 501
Private Const P_NOT As Long = 502
Private Const P_EQP As Long = 503
Private Const P_EQVP As Long = 504
Private Const P_EQUALP As Long = 505
Private Const P_DISPLAY As Long = 506
Private Const P_WRITE As Long = 507
Private Const P_NEWLINE As Long = 508
Private Const P_WRITESTR As Long = 509
Private Const P_WRITECHAR As Long = 510
Private Const P_DISPLAYLN As Long = 511
Private Const P_FORMAT As Long = 512
Private Const P_ERROR As Long = 513
Private Const P_RAISE As Long = 514
Private Const P_WITHHANDLER As Long = 515
Private Const P_ERROBJP As Long = 516
Private Const P_ERRMSG As Long = 517
Private Const P_ERRIRR As Long = 518
Private Const P_REPORTSTR As Long = 519
Private Const P_EXIT As Long = 520
Private Const P_VOID As Long = 521
Private Const P_FORCE As Long = 522
Private Const P_MAKEPROMISE As Long = 523
Private Const P_PROMISEP As Long = 524
Private Const P_CALLCC As Long = 525
Private Const P_VALUES As Long = 526
Private Const P_CALLWV As Long = 527
Private Const P_DYNWIND As Long = 528
Private Const P_EVAL As Long = 529
Private Const P_INTERENV As Long = 530
Private Const P_RUNTIME As Long = 531
Private Const P_STREAMCAR As Long = 532
Private Const P_STREAMCDR As Long = 533
Private Const P_STREAMPAIRP As Long = 534
Private Const P_READLINE As Long = 535
Private Const P_READ As Long = 536
Private Const P_IDENTITY As Long = 537
Private Const P_SYMEQ As Long = 538
Private Const P_BOOLEQ As Long = 539
Private Const P_WRITELINE As Long = 540
Private Const P_PROCARITY As Long = 541
Private Const P_CURSECONDS As Long = 542
' Excel 600..699 (implemented in LispExcel)
Public Const LP_X_CELL As Long = 600
Public Const LP_X_SETCELL As Long = 601
Public Const LP_X_GETRANGE As Long = 602
Public Const LP_X_SETRANGE As Long = 603
Public Const LP_X_CLEAR As Long = 604
Public Const LP_X_MSGBOX As Long = 605
Public Const LP_X_EVAL As Long = 606
Public Const LP_X_SHEETS As Long = 607

Private mRndInit As Boolean

' ---------------------------------------------------------------------------
'  registration
' ---------------------------------------------------------------------------
Private Sub Reg(ByVal nm As String, ByVal code As Long, ByVal mn As Long, ByVal mx As Long)
    Dim o As LispObj, s As LispObj, v As Variant
    Set o = LNewObj(LK_PRIM)
    o.SName = nm
    o.Code = code
    o.I1 = mn
    o.I2 = mx
    Set s = LIntern(nm)
    Set v = o
    lgGlobal.SetRef s, s.Id, v
End Sub

Private Sub RegVar(ByVal nm As String, ByRef v As Variant)
    Dim s As LispObj
    Set s = LIntern(nm)
    lgGlobal.SetRef s, s.Id, v
End Sub

Public Sub LInitBuiltins()
    Dim i As Long, nm As String, v As Variant, names As Variant
    ' numbers
    Reg "+", P_ADD, 0, -1
    Reg "-", P_SUB, 1, -1
    Reg "*", P_MUL, 0, -1
    Reg "/", P_DIV, 1, -1
    Reg "=", P_NUMEQ, 1, -1
    Reg "<", P_LT, 1, -1
    Reg ">", P_GT, 1, -1
    Reg "<=", P_LE, 1, -1
    Reg ">=", P_GE, 1, -1
    Reg "quotient", P_QUOTIENT, 2, 2
    Reg "truncate-quotient", P_QUOTIENT, 2, 2
    Reg "remainder", P_REMAINDER, 2, 2
    Reg "truncate-remainder", P_REMAINDER, 2, 2
    Reg "modulo", P_MODULO, 2, 2
    Reg "floor-remainder", P_MODULO, 2, 2
    Reg "abs", P_ABS, 1, 1
    Reg "magnitude", P_ABS, 1, 1
    Reg "min", P_MIN, 1, -1
    Reg "max", P_MAX, 1, -1
    Reg "gcd", P_GCD, 0, -1
    Reg "lcm", P_LCM, 0, -1
    Reg "expt", P_EXPT, 2, 2
    Reg "exp", P_EXP, 1, 1
    Reg "log", P_LOG, 1, 2
    Reg "sin", P_SIN, 1, 1
    Reg "cos", P_COS, 1, 1
    Reg "tan", P_TAN, 1, 1
    Reg "asin", P_ASIN, 1, 1
    Reg "acos", P_ACOS, 1, 1
    Reg "atan", P_ATAN, 1, 2
    Reg "sqrt", P_SQRT, 1, 1
    Reg "exact-integer-sqrt", P_EXACTSQRT, 1, 1
    Reg "square", P_SQUARE, 1, 1
    Reg "number?", P_NUMBERP, 1, 1
    Reg "complex?", P_NUMBERP, 1, 1
    Reg "real?", P_NUMBERP, 1, 1
    Reg "integer?", P_INTEGERP, 1, 1
    Reg "rational?", P_RATIONALP, 1, 1
    Reg "exact?", P_EXACTP, 1, 1
    Reg "inexact?", P_INEXACTP, 1, 1
    Reg "exact-integer?", P_EXACTINTP, 1, 1
    Reg "exact-rational?", P_EXACTINTP, 1, 1
    Reg "zero?", P_ZEROP, 1, 1
    Reg "positive?", P_POSITIVEP, 1, 1
    Reg "negative?", P_NEGATIVEP, 1, 1
    Reg "odd?", P_ODDP, 1, 1
    Reg "even?", P_EVENP, 1, 1
    Reg "nan?", P_NANP, 1, 1
    Reg "exact->inexact", P_INEXACT, 1, 1
    Reg "inexact", P_INEXACT, 1, 1
    Reg "inexact->exact", P_EXACT, 1, 1
    Reg "exact", P_EXACT, 1, 1
    Reg "floor", P_FLOOR, 1, 1
    Reg "ceiling", P_CEILING, 1, 1
    Reg "round", P_ROUND, 1, 1
    Reg "truncate", P_TRUNCATE, 1, 1
    Reg "number->string", P_NUM2STR, 1, 2
    Reg "string->number", P_STR2NUM, 1, 2
    Reg "random", P_RANDOM, 1, 1
    Reg "1+", P_ADD1, 1, 1
    Reg "add1", P_ADD1, 1, 1
    Reg "1-", P_SUB1, 1, 1
    Reg "-1+", P_SUB1, 1, 1
    Reg "sub1", P_SUB1, 1, 1
    ' pairs and lists
    Reg "cons", P_CONS, 2, 2
    Reg "car", P_CAR, 1, 1
    Reg "cdr", P_CDR, 1, 1
    Reg "set-car!", P_SETCAR, 2, 2
    Reg "set-cdr!", P_SETCDR, 2, 2
    names = Array("caar", "cadr", "cdar", "cddr", "caaar", "caadr", "cadar", "caddr", "cdaar", "cdadr", "cddar", "cdddr", "cadddr", "cddddr")
    For i = 0 To UBound(names)
        Reg CStr(names(i)), P_CXR, 1, 1
    Next
    Reg "first", P_CXR, 1, 1
    Reg "second", P_CXR, 1, 1
    Reg "third", P_CXR, 1, 1
    Reg "fourth", P_CXR, 1, 1
    Reg "fifth", P_CXR, 1, 1
    Reg "rest", P_CDR, 1, 1
    Reg "list", P_LIST, 0, -1
    Reg "cons*", P_CONSSTAR, 1, -1
    Reg "list*", P_CONSSTAR, 1, -1
    Reg "length", P_LENGTH, 1, 1
    Reg "append", P_APPEND, 0, -1
    Reg "reverse", P_REVERSE, 1, 1
    Reg "list-tail", P_LISTTAIL, 2, 2
    Reg "list-ref", P_LISTREF, 2, 2
    Reg "list-copy", P_LISTCOPY, 1, 1
    Reg "last-pair", P_LASTPAIR, 1, 1
    Reg "last", P_LAST, 1, 1
    Reg "memq", P_MEMQ, 2, 2
    Reg "memv", P_MEMV, 2, 2
    Reg "member", P_MEMBER, 2, 3
    Reg "assq", P_ASSQ, 2, 2
    Reg "assv", P_ASSV, 2, 2
    Reg "assoc", P_ASSOC, 2, 3
    Reg "map", P_MAP, 2, -1
    Reg "for-each", P_FOREACH, 2, -1
    Reg "filter", P_FILTER, 2, 2
    Reg "remove", P_REMOVE, 2, 2
    Reg "reduce", P_REDUCE, 3, 3
    Reg "fold-left", P_FOLDLEFT, 3, -1
    Reg "fold-right", P_FOLDRIGHT, 3, -1
    Reg "fold", P_FOLD, 3, -1
    Reg "foldl", P_FOLD, 3, -1
    Reg "foldr", P_FOLDRIGHT, 3, -1
    Reg "delete", P_DELETE, 2, 2
    Reg "list-index", P_LISTINDEX, 2, 2
    Reg "iota", P_IOTA, 1, 3
    Reg "apply", LP_APPLY, 1, -1
    Reg "sort", P_SORT, 2, 2
    Reg "sort!", P_SORT, 2, 2
    Reg "list-sort", P_LISTSORT, 2, 2
    Reg "any", P_ANY, 2, -1
    Reg "every", P_EVERY, 2, -1
    Reg "count", P_COUNT, 2, 2
    Reg "delete-duplicates", P_DELDUP, 1, 2
    Reg "remove-duplicates", P_DELDUP, 1, 2
    Reg "append-map", P_APPENDMAP, 2, 2
    Reg "find", P_FIND, 2, 2
    Reg "list?", P_LISTP, 1, 1
    Reg "pair?", P_PAIRP, 1, 1
    Reg "null?", P_NULLP, 1, 1
    Reg "empty?", P_NULLP, 1, 1
    Reg "make-list", P_MAKELIST, 1, 2
    Reg "list-set!", P_LISTSET, 3, 3
    Reg "filter-map", P_FILTERMAP, 2, 2
    Reg "partition", P_PARTITION, 2, 2
    ' strings, characters, symbols
    Reg "symbol?", P_SYMBOLP, 1, 1
    Reg "string?", P_STRINGP, 1, 1
    Reg "char?", P_CHARP, 1, 1
    Reg "symbol->string", P_SYM2STR, 1, 1
    Reg "string->symbol", P_STR2SYM, 1, 1
    Reg "intern", P_STR2SYM, 1, 1
    Reg "gensym", P_GENSYM, 0, 1
    Reg "generate-uninterned-symbol", P_GENSYM, 0, 1
    Reg "string-length", P_STRLEN, 1, 1
    Reg "string-ref", P_STRREF, 2, 2
    Reg "substring", P_SUBSTRING, 2, 3
    Reg "string-append", P_STRAPPEND, 0, -1
    Reg "string-copy", P_STRCOPY, 1, 3
    Reg "string=?", P_STREQ, 1, -1
    Reg "string<?", P_STRLT, 1, -1
    Reg "string>?", P_STRGT, 1, -1
    Reg "string<=?", P_STRLE, 1, -1
    Reg "string>=?", P_STRGE, 1, -1
    Reg "string-ci=?", P_STRCIEQ, 1, -1
    Reg "string-upcase", P_STRUP, 1, 1
    Reg "string-downcase", P_STRDOWN, 1, 1
    Reg "string->list", P_STR2LIST, 1, 3
    Reg "list->string", P_LIST2STR, 1, 1
    Reg "string", P_STRING, 0, -1
    Reg "make-string", P_MAKESTR, 1, 2
    Reg "string-null?", P_STRNULLP, 1, 1
    Reg "string-index", P_STRINDEX, 2, 2
    Reg "string-contains", P_STRCONTAINS, 2, 2
    Reg "string-search-forward", P_STRCONTAINS, 2, 3
    Reg "string-contains?", P_STRCONTAINSP, 2, 2
    Reg "string-split", P_STRSPLIT, 1, 2
    Reg "string-join", P_STRJOIN, 1, 2
    Reg "string-trim", P_STRTRIM, 1, 1
    Reg "string-trim-both", P_STRTRIM, 1, 1
    Reg "string-trim-left", P_STRTRIML, 1, 1
    Reg "string-trim-right", P_STRTRIMR, 1, 1
    Reg "string-reverse", P_STRREV, 1, 1
    Reg "string-prefix?", P_STRPREFIXP, 2, 2
    Reg "string-suffix?", P_STRSUFFIXP, 2, 2
    Reg "string-for-each", P_STRFOREACH, 2, 2
    Reg "string-map", P_STRMAP, 2, 2
    Reg "string-pad", P_STRPAD, 2, 3
    Reg "string-pad-left", P_STRPAD, 2, 3
    Reg "string-pad-right", P_STRPADR, 2, 3
    Reg "string-replace-all", P_STRREPLACE, 3, 3
    Reg "char->integer", P_CHAR2INT, 1, 1
    Reg "integer->char", P_INT2CHAR, 1, 1
    Reg "char-upcase", P_CHARUP, 1, 1
    Reg "char-downcase", P_CHARDOWN, 1, 1
    Reg "char-alphabetic?", P_CHARALPHAP, 1, 1
    Reg "char-numeric?", P_CHARNUMP, 1, 1
    Reg "char-whitespace?", P_CHARWSP, 1, 1
    Reg "char-upper-case?", P_CHARUPPERP, 1, 1
    Reg "char-lower-case?", P_CHARLOWERP, 1, 1
    Reg "digit-value", P_DIGITVAL, 1, 1
    Reg "char->digit", P_DIGITVAL, 1, 2
    Reg "char=?", P_CHAREQ, 1, -1
    Reg "char<?", P_CHARLT, 1, -1
    Reg "char>?", P_CHARGT, 1, -1
    Reg "char<=?", P_CHARLE, 1, -1
    Reg "char>=?", P_CHARGE, 1, -1
    ' vectors and hash tables
    Reg "vector?", P_VECTORP, 1, 1
    Reg "make-vector", P_MAKEVECTOR, 1, 2
    Reg "vector", P_VECTOR, 0, -1
    Reg "vector-ref", P_VECREF, 2, 2
    Reg "vector-set!", P_VECSET, 3, 3
    Reg "vector-length", P_VECLEN, 1, 1
    Reg "vector->list", P_VEC2LIST, 1, 3
    Reg "list->vector", P_LIST2VEC, 1, 1
    Reg "vector-fill!", P_VECFILL, 2, 2
    Reg "vector-map", P_VECMAP, 2, -1
    Reg "vector-for-each", P_VECFOREACH, 2, -1
    Reg "vector-copy", P_VECCOPY, 1, 3
    names = Array("make-hash-table", "make-equal-hash-table", "make-strong-eqv-hash-table", "make-string-hash-table", "make-eqv-hash-table", "make-eq-hash-table", "make-hash")
    For i = 0 To UBound(names)
        Reg CStr(names(i)), P_MAKEHASH, 0, -1
    Next
    Reg "hash-table?", P_HASHP, 1, 1
    Reg "hash-table-set!", P_HASHSET, 3, 3
    Reg "hash-table/put!", P_HASHSET, 3, 3
    Reg "hash-set!", P_HASHSET, 3, 3
    Reg "hash-table-ref", P_HASHREF, 2, 3
    Reg "hash-table-ref/default", P_HASHREFDEF, 3, 3
    Reg "hash-table/get", P_HASHREFDEF, 3, 3
    Reg "hash-ref", P_HASHGUILEREF, 2, 3
    Reg "hash-table-delete!", P_HASHDEL, 2, 2
    Reg "hash-table/remove!", P_HASHDEL, 2, 2
    Reg "hash-remove!", P_HASHDEL, 2, 2
    Reg "hash-table-contains?", P_HASHHAS, 2, 2
    Reg "hash-table-exists?", P_HASHHAS, 2, 2
    Reg "hash-table-count", P_HASHCOUNT, 1, 1
    Reg "hash-table-size", P_HASHCOUNT, 1, 1
    Reg "hash-table/count", P_HASHCOUNT, 1, 1
    Reg "hash-table-keys", P_HASHKEYS, 1, 1
    Reg "hash-table/key-list", P_HASHKEYS, 1, 1
    Reg "hash-table-values", P_HASHVALUES, 1, 1
    Reg "hash-table/datum-list", P_HASHVALUES, 1, 1
    Reg "hash-table->alist", P_HASH2ALIST, 1, 1
    Reg "hash-table-update!", P_HASHUPDATE, 3, 4
    Reg "hash-table-update!/default", P_HASHUPDATEDEF, 4, 4
    Reg "hash-table-walk", P_HASHWALK, 2, 2
    Reg "hash-table-clear!", P_HASHCLEAR, 1, 1
    Reg "hash-table-copy", P_HASHCOPY, 1, 2
    Reg "hash-for-each", P_HASHFOREACH, 2, 2
    ' control and output
    Reg "procedure?", P_PROCP, 1, 1
    Reg "boolean?", P_BOOLP, 1, 1
    Reg "not", P_NOT, 1, 1
    Reg "eq?", P_EQP, 2, 2
    Reg "eqv?", P_EQVP, 2, 2
    Reg "equal?", P_EQUALP, 2, 2
    Reg "symbol=?", P_SYMEQ, 1, -1
    Reg "boolean=?", P_BOOLEQ, 1, -1
    Reg "display", P_DISPLAY, 1, 2
    Reg "write", P_WRITE, 1, 2
    Reg "write-shared", P_WRITE, 1, 2
    Reg "write-simple", P_WRITE, 1, 2
    Reg "newline", P_NEWLINE, 0, 1
    Reg "write-string", P_WRITESTR, 1, 2
    Reg "write-char", P_WRITECHAR, 1, 2
    Reg "displayln", P_DISPLAYLN, 1, 2
    Reg "write-line", P_WRITELINE, 1, 2
    Reg "format", P_FORMAT, 1, -1
    Reg "error", P_ERROR, 0, -1
    Reg "raise", P_RAISE, 1, 1
    Reg "raise-continuable", P_RAISE, 1, 1
    Reg "with-exception-handler", P_WITHHANDLER, 2, 2
    Reg "error-object?", P_ERROBJP, 1, 1
    Reg "error?", P_ERROBJP, 1, 1
    Reg "condition?", P_ERROBJP, 1, 1
    Reg "error-object-message", P_ERRMSG, 1, 1
    Reg "error-message", P_ERRMSG, 1, 1
    Reg "error-object-irritants", P_ERRIRR, 1, 1
    Reg "error-irritants", P_ERRIRR, 1, 1
    Reg "condition/report-string", P_REPORTSTR, 1, 1
    Reg "exit", P_EXIT, 0, 1
    Reg "void", P_VOID, 0, -1
    Reg "force", P_FORCE, 1, 1
    Reg "make-promise", P_MAKEPROMISE, 1, 1
    Reg "promise?", P_PROMISEP, 1, 1
    Reg "call/cc", P_CALLCC, 1, 1
    Reg "call-with-current-continuation", P_CALLCC, 1, 1
    Reg "call-with-escape-continuation", P_CALLCC, 1, 1
    Reg "values", P_VALUES, 0, -1
    Reg "call-with-values", P_CALLWV, 2, 2
    Reg "dynamic-wind", P_DYNWIND, 3, 3
    Reg "eval", P_EVAL, 1, 2
    Reg "interaction-environment", P_INTERENV, 0, 0
    Reg "the-environment", P_INTERENV, 0, 0
    Reg "system-global-environment", P_INTERENV, 0, 0
    Reg "runtime", P_RUNTIME, 0, 0
    Reg "current-milliseconds", P_RUNTIME, 0, 0
    Reg "current-time", P_CURSECONDS, 0, 0
    Reg "current-seconds", P_CURSECONDS, 0, 0
    Reg "stream-car", P_STREAMCAR, 1, 1
    Reg "stream-cdr", P_STREAMCDR, 1, 1
    Reg "stream-pair?", P_STREAMPAIRP, 1, 1
    Reg "stream-null?", P_NULLP, 1, 1
    Reg "empty-stream?", P_NULLP, 1, 1
    Reg "read-line", P_READLINE, 0, 1
    Reg "read", P_READ, 0, 1
    Reg "identity", P_IDENTITY, 1, 1
    Reg "values-identity", P_IDENTITY, 1, 1
    Reg "procedure-arity", P_PROCARITY, 1, 1
    v = Empty
    RegVar "the-empty-stream", v
    RegVar "stream-nil", v
    RegVar "empty", v
    RegVar "null", v
    v = True
    RegVar "true", v
    v = False
    RegVar "false", v
    ' Excel
    Reg "cell", LP_X_CELL, 1, 3
    Reg "set-cell!", LP_X_SETCELL, 2, 4
    Reg "get-range", LP_X_GETRANGE, 1, 2
    Reg "set-range!", LP_X_SETRANGE, 2, 3
    Reg "clear-range!", LP_X_CLEAR, 1, 2
    Reg "msgbox", LP_X_MSGBOX, 1, 2
    Reg "excel-eval", LP_X_EVAL, 1, 1
    Reg "sheet-names", LP_X_SHEETS, 0, 0
End Sub

' ---------------------------------------------------------------------------
'  argument helpers
' ---------------------------------------------------------------------------
Private Function TypeErr(ByVal who As String, ByVal what As String, ByRef v As Variant) As Boolean
    LThrow2 who & ": expected " & what & ", got", v
End Function

Private Function NumArg(ByRef v As Variant, ByVal who As String) As Variant
    If Not LIsNum(v) Then TypeErr who, "a number", v
    If VarType(v) = vbInteger Or VarType(v) = vbByte Then
        NumArg = CLng(v)
    Else
        NumArg = v
    End If
End Function

Private Function DblArg(ByRef v As Variant, ByVal who As String) As Double
    If Not LIsNum(v) Then TypeErr who, "a number", v
    DblArg = CDbl(v)
End Function

Private Function IntArg(ByRef v As Variant, ByVal who As String) As Variant
    If LIsExact(v) Then
        IntArg = LNorm(v)
    ElseIf LIsFlo(v) Then
        If CDbl(v) <> Int(CDbl(v)) Then TypeErr who, "an integer", v
        IntArg = LIntFromDouble(CDbl(v))
    Else
        TypeErr who, "an integer", v
    End If
End Function

Private Function LongArg(ByRef v As Variant, ByVal who As String) As Long
    Dim x As Variant
    x = IntArg(v, who)
    If VarType(x) = vbDecimal Then LThrow2 who & ": integer too large", v
    LongArg = x
End Function

Private Function StrArg(ByRef v As Variant, ByVal who As String) As String
    If VarType(v) <> vbString Then TypeErr who, "a string", v
    StrArg = v
End Function

Private Function CharArg(ByRef v As Variant, ByVal who As String) As String
    Dim o As LispObj
    If Not LIsKind(v, LK_CHAR) Then TypeErr who, "a character", v
    Set o = v
    CharArg = o.SName
End Function

Private Function PairArg(ByRef v As Variant, ByVal who As String) As LispObj
    If Not LIsPair(v) Then TypeErr who, "a pair", v
    Set PairArg = v
End Function

Private Function ObjArg(ByRef v As Variant, ByVal kind As Long, ByVal who As String, ByVal what As String) As LispObj
    If Not LIsKind(v, kind) Then TypeErr who, what, v
    Set ObjArg = v
End Function

Private Sub ProcArg(ByRef v As Variant, ByVal who As String)
    If Not LIsProc(v) Then TypeErr who, "a procedure", v
End Sub

Private Function ToV(ByVal o As LispObj) As Variant
    Set ToV = o
End Function

' list builder
Private Sub LB_Add(ByRef head As LispObj, ByRef tail As LispObj, ByRef v As Variant)
    Dim cell As LispObj, e As Variant
    Set cell = LCons(v, e)
    If head Is Nothing Then
        Set head = cell
    Else
        Set tail.Cdr = cell
    End If
    Set tail = cell
End Sub

Private Sub LB_Result(ByVal head As LispObj, ByRef out As Variant)
    If head Is Nothing Then
        out = Empty
    Else
        Set out = head
    End If
End Sub

Private Function MakeValues(ByRef args() As Variant, ByVal na As Long) As Variant
    Dim o As LispObj, i As Long
    If na = 1 Then
        If IsObject(args(0)) Then Set MakeValues = args(0) Else MakeValues = args(0)
        Exit Function
    End If
    Set o = LNewObj(LK_VALUES)
    For i = 0 To na - 1
        o.Push args(i)
    Next
    Set MakeValues = o
End Function

' ---------------------------------------------------------------------------
'  dispatch
' ---------------------------------------------------------------------------
Public Sub LCallPrim(ByVal fo As LispObj, ByRef args() As Variant, ByVal na As Long, ByRef out As Variant)
    Dim c As Long, want As String
    If fo.Kind = LK_RECPROC Then
        RecordProc fo, args, na, out
        Exit Sub
    End If
    If na < fo.I1 Or (fo.I2 >= 0 And na > fo.I2) Then
        If fo.I2 = fo.I1 Then
            want = CStr(fo.I1)
        ElseIf fo.I2 < 0 Then
            want = "at least " & fo.I1
        Else
            want = fo.I1 & " to " & fo.I2
        End If
        LThrow fo.SName & ": wrong number of arguments (expected " & want & ", got " & na & ")"
    End If
    c = fo.Code
    If c < 200 Then
        PrimNum fo, c, args, na, out
    ElseIf c < 300 Then
        PrimList fo, c, args, na, out
    ElseIf c < 400 Then
        PrimStr fo, c, args, na, out
    ElseIf c < 500 Then
        PrimVec fo, c, args, na, out
    ElseIf c < 600 Then
        PrimCtl fo, c, args, na, out
    Else
        LispHostCall c, args, na, out
    End If
End Sub

Private Sub RecordProc(ByVal fo As LispObj, ByRef args() As Variant, ByVal na As Long, ByRef out As Variant)
    Dim r As LispObj, i As Long, idx As Variant, e As Variant, rt As LispObj
    Set rt = fo.Ref
    Select Case fo.Code
    Case 1
        If na <> fo.I1 Then LThrow fo.SName & ": wrong number of arguments (expected " & fo.I1 & ", got " & na & ")"
        Set r = LNewObj(LK_RECORD)
        Set r.Ref = rt
        For i = 0 To rt.Count - 1
            r.Push e
        Next
        For i = 0 To na - 1
            fo.GetInto i, idx
            r.SetI CLng(idx), args(i)
        Next
        Set out = r
    Case 2
        If na <> 1 Then LThrow fo.SName & ": expected 1 argument"
        out = False
        If LIsKind(args(0), LK_RECORD) Then
            Set r = args(0)
            out = (r.Ref Is rt)
        End If
    Case 3, 4
        If na <> fo.Code - 2 Then LThrow fo.SName & ": wrong number of arguments"
        If Not LIsKind(args(0), LK_RECORD) Then TypeErr fo.SName, "a " & rt.SName & " record", args(0)
        Set r = args(0)
        If Not (r.Ref Is rt) Then TypeErr fo.SName, "a " & rt.SName & " record", args(0)
        If fo.Code = 3 Then
            r.GetInto fo.I1, out
        Else
            r.SetI fo.I1, args(1)
            Set out = lgVoid
        End If
    End Select
End Sub

' ---------------------------------------------------------------------------
'  numbers
' ---------------------------------------------------------------------------
Private Function NumAdd(ByRef a As Variant, ByRef b As Variant) As Variant
    Dim d As Double
    If VarType(a) = vbLong And VarType(b) = vbLong Then
        d = CDbl(a) + CDbl(b)
        If d >= -2147483648# And d <= 2147483647# Then
            NumAdd = CLng(d)
            Exit Function
        End If
    End If
    If LIsExact(a) And LIsExact(b) Then
        NumAdd = LIntAdd(a, b)
    Else
        NumAdd = CDbl(a) + CDbl(b)
    End If
End Function

Private Function NumSub(ByRef a As Variant, ByRef b As Variant) As Variant
    If LIsExact(a) And LIsExact(b) Then
        NumSub = LIntSub(a, b)
    Else
        NumSub = CDbl(a) - CDbl(b)
    End If
End Function

Private Function NumMul(ByRef a As Variant, ByRef b As Variant) As Variant
    If LIsExact(a) And LIsExact(b) Then
        NumMul = LIntMul(a, b)
    Else
        NumMul = CDbl(a) * CDbl(b)
    End If
End Function

Private Function NumDiv(ByRef a As Variant, ByRef b As Variant) As Variant
    If LIsExact(a) And LIsExact(b) Then
        NumDiv = LExactDiv(a, b)
    Else
        If CDbl(b) = 0 Then LThrow "/: division by zero"
        NumDiv = CDbl(a) / CDbl(b)
    End If
End Function

Private Function IsPerfectSquare(ByVal n As Variant, ByRef root As Variant) As Boolean
    Dim r As Variant
    r = LIntFromDouble(Int(Sqr(CDbl(n))))
    Do While LIntMul(r, r) > n
        r = LIntSub(r, 1&)
    Loop
    Do While LIntMul(LIntAdd(r, 1&), LIntAdd(r, 1&)) <= n
        r = LIntAdd(r, 1&)
    Loop
    root = r
    IsPerfectSquare = (LIntMul(r, r) = n)
End Function

Private Function RandDbl() As Double
    If Not mRndInit Then
        Randomize
        mRndInit = True
    End If
    RandDbl = (Int(Rnd() * 65536#) * 65536# + Int(Rnd() * 65536#)) / 4294967296#
End Function

Private Function GcdV(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim q As Variant, r As Variant
    If a < 0 Then a = LIntNeg(a)
    If b < 0 Then b = LIntNeg(b)
    Do While b <> 0
        LIntDivMod a, b, q, r
        a = b
        b = r
    Loop
    GcdV = a
End Function

Private Function IntToRadix(ByVal x As Variant, ByVal bs As Long) As String
    Dim q As Variant, r As Variant, s As String, neg As Boolean
    If x = 0 Then
        IntToRadix = "0"
        Exit Function
    End If
    If x < 0 Then
        neg = True
        x = LIntNeg(x)
    End If
    Do While x > 0
        LIntDivMod x, bs, q, r
        s = Mid$("0123456789abcdefghijklmnopqrstuvwxyz", r + 1, 1) & s
        x = q
    Loop
    If neg Then s = "-" & s
    IntToRadix = s
End Function

Private Sub PrimNum(ByVal fo As LispObj, ByVal c As Long, ByRef args() As Variant, ByVal na As Long, ByRef out As Variant)
    Dim acc As Variant, i As Long, x As Variant, y As Variant, d As Double, d2 As Double, q As Variant, r As Variant
    Dim who As String, ok As Boolean, inexact As Boolean, cmp As Long, bs As Long
    who = fo.SName
    Select Case c
    Case P_ADD
        acc = 0&
        For i = 0 To na - 1
            acc = NumAdd(acc, NumArg(args(i), who))
        Next
        out = acc
    Case P_MUL
        acc = 1&
        For i = 0 To na - 1
            acc = NumMul(acc, NumArg(args(i), who))
        Next
        out = acc
    Case P_SUB
        acc = NumArg(args(0), who)
        If na = 1 Then
            If LIsExact(acc) Then
                out = LIntNeg(acc)
            Else
                out = -CDbl(acc)
            End If
            Exit Sub
        End If
        For i = 1 To na - 1
            acc = NumSub(acc, NumArg(args(i), who))
        Next
        out = acc
    Case P_DIV
        acc = NumArg(args(0), who)
        If na = 1 Then
            out = NumDiv(1&, acc)
            Exit Sub
        End If
        For i = 1 To na - 1
            acc = NumDiv(acc, NumArg(args(i), who))
        Next
        out = acc
    Case P_NUMEQ, P_LT, P_GT, P_LE, P_GE
        out = True
        x = NumArg(args(0), who)
        For i = 1 To na - 1
            y = NumArg(args(i), who)
            cmp = LNumCmp(x, y)
            Select Case c
            Case P_NUMEQ: ok = (cmp = 0)
            Case P_LT: ok = (cmp < 0)
            Case P_GT: ok = (cmp > 0)
            Case P_LE: ok = (cmp <= 0)
            Case Else: ok = (cmp >= 0)
            End Select
            If Not ok Then
                out = False
                Exit Sub
            End If
            x = y
        Next
    Case P_QUOTIENT, P_REMAINDER, P_MODULO
        x = NumArg(args(0), who)
        y = NumArg(args(1), who)
        If LIsExact(x) And LIsExact(y) Then
            If y = 0 Then LThrow who & ": division by zero"
            If c = P_MODULO Then
                LIntDivMod x, y, q, r
            Else
                LIntQuotRem x, y, q, r
            End If
            If c = P_QUOTIENT Then
                out = q
            Else
                out = r
            End If
        Else
            d = CDbl(x)
            d2 = CDbl(y)
            If d2 = 0 Then LThrow who & ": division by zero"
            Select Case c
            Case P_QUOTIENT: out = Fix(d / d2)
            Case P_REMAINDER: out = d - d2 * Fix(d / d2)
            Case Else: out = d - d2 * Int(d / d2)
            End Select
        End If
    Case P_ABS
        x = NumArg(args(0), who)
        If LIsExact(x) Then
            If x < 0 Then
                out = LIntNeg(x)
            Else
                out = x
            End If
        Else
            out = Abs(CDbl(x))
        End If
    Case P_MIN, P_MAX
        acc = NumArg(args(0), who)
        inexact = LIsFlo(acc)
        For i = 1 To na - 1
            x = NumArg(args(i), who)
            If LIsFlo(x) Then inexact = True
            If c = P_MIN Then
                If LNumCmp(x, acc) < 0 Then acc = x
            Else
                If LNumCmp(x, acc) > 0 Then acc = x
            End If
        Next
        If inexact Then
            out = CDbl(acc)
        Else
            out = acc
        End If
    Case P_GCD
        acc = 0&
        For i = 0 To na - 1
            acc = GcdV(acc, IntArg(args(i), who))
        Next
        out = acc
    Case P_LCM
        acc = 1&
        For i = 0 To na - 1
            x = IntArg(args(i), who)
            If x < 0 Then x = LIntNeg(x)
            If x = 0 Or acc = 0 Then
                acc = 0&
            Else
                LIntDivMod LIntMul(acc, x), GcdV(acc, x), q, r
                acc = q
            End If
        Next
        out = acc
    Case P_EXPT
        x = NumArg(args(0), who)
        y = NumArg(args(1), who)
        If LIsExact(x) And LIsExact(y) Then
            If y >= 0 Then
                out = LIntPow(x, y)
            Else
                out = NumDiv(1&, LIntPow(x, LIntNeg(y)))
            End If
        Else
            d = CDbl(x)
            d2 = CDbl(y)
            If d < 0 And d2 <> Int(d2) Then LThrow "expt: complex results are not supported"
            If d = 0 And d2 < 0 Then LThrow "expt: division by zero"
            out = d ^ d2
        End If
    Case P_EXP
        out = Exp(DblArg(args(0), who))
    Case P_LOG
        d = DblArg(args(0), who)
        If d <= 0 Then LThrow2 "log: argument must be positive, got", args(0)
        If na = 2 Then
            d2 = DblArg(args(1), who)
            out = Log(d) / Log(d2)
        Else
            out = Log(d)
        End If
    Case P_SIN
        out = Sin(DblArg(args(0), who))
    Case P_COS
        out = Cos(DblArg(args(0), who))
    Case P_TAN
        out = Tan(DblArg(args(0), who))
    Case P_ASIN, P_ACOS
        d = DblArg(args(0), who)
        If d < -1 Or d > 1 Then LThrow2 who & ": argument out of range", args(0)
        If Abs(d) = 1 Then
            d2 = Sgn(d) * 2 * Atn(1#)
        Else
            d2 = Atn(d / Sqr(1 - d * d))
        End If
        If c = P_ASIN Then
            out = d2
        Else
            out = 2 * Atn(1#) - d2
        End If
    Case P_ATAN
        If na = 1 Then
            out = Atn(DblArg(args(0), who))
        Else
            d = DblArg(args(0), who)
            d2 = DblArg(args(1), who)
            If d2 > 0 Then
                out = Atn(d / d2)
            ElseIf d2 < 0 Then
                If d >= 0 Then
                    out = Atn(d / d2) + 4 * Atn(1#)
                Else
                    out = Atn(d / d2) - 4 * Atn(1#)
                End If
            ElseIf d > 0 Then
                out = 2 * Atn(1#)
            ElseIf d < 0 Then
                out = -2 * Atn(1#)
            Else
                out = 0#
            End If
        End If
    Case P_SQRT
        x = NumArg(args(0), who)
        If CDbl(x) < 0 Then LThrow "sqrt: complex results are not supported"
        If LIsExact(x) Then
            If IsPerfectSquare(x, r) Then
                out = r
                Exit Sub
            End If
        End If
        out = Sqr(CDbl(x))
    Case P_EXACTSQRT
        x = IntArg(args(0), who)
        IsPerfectSquare x, r
        LSetV out, MakeValues2(r, LIntSub(x, LIntMul(r, r)))
    Case P_SQUARE
        x = NumArg(args(0), who)
        out = NumMul(x, x)
    Case P_NUMBERP
        out = LIsNum(args(0))
    Case P_INTEGERP
        If LIsExact(args(0)) Then
            out = True
        ElseIf LIsFlo(args(0)) Then
            out = (CDbl(args(0)) = Int(CDbl(args(0))) And Abs(CDbl(args(0))) < L_INF)
        Else
            out = False
        End If
    Case P_RATIONALP
        If LIsNum(args(0)) Then
            out = (Abs(CDbl(args(0))) < L_INF)
        Else
            out = False
        End If
    Case P_EXACTP
        out = LIsExact(NumArg(args(0), who))
    Case P_INEXACTP
        out = LIsFlo(NumArg(args(0), who))
    Case P_EXACTINTP
        out = LIsExact(args(0))
    Case P_NANP
        NumArg args(0), who
        out = False
    Case P_ZEROP
        out = (CDbl(NumArg(args(0), who)) = 0)
    Case P_POSITIVEP
        out = (CDbl(NumArg(args(0), who)) > 0)
    Case P_NEGATIVEP
        out = (CDbl(NumArg(args(0), who)) < 0)
    Case P_ODDP, P_EVENP
        x = IntArg(args(0), who)
        LIntDivMod x, 2&, q, r
        If c = P_ODDP Then
            out = (r <> 0)
        Else
            out = (r = 0)
        End If
    Case P_INEXACT
        out = CDbl(NumArg(args(0), who))
    Case P_EXACT
        x = NumArg(args(0), who)
        If LIsExact(x) Then
            out = x
        Else
            d = CDbl(x)
            If d <> Int(d) Then LThrow2 "exact: no exact integer for", x
            out = LIntFromDouble(d)
        End If
    Case P_FLOOR, P_CEILING, P_ROUND, P_TRUNCATE
        x = NumArg(args(0), who)
        If LIsExact(x) Then
            out = x
        Else
            d = CDbl(x)
            Select Case c
            Case P_FLOOR: out = Int(d)
            Case P_CEILING: out = -Int(-d)
            Case P_ROUND: out = CDbl(Round(d))
            Case Else: out = Fix(d)
            End Select
        End If
    Case P_NUM2STR
        x = NumArg(args(0), who)
        bs = 10
        If na = 2 Then bs = LongArg(args(1), who)
        If bs <> 10 And LIsExact(x) Then
            out = IntToRadix(x, bs)
        Else
            out = LToStr(x, False)
        End If
    Case P_STR2NUM
        bs = 10
        If na = 2 Then bs = LongArg(args(1), who)
        x = LParseNumber(Trim$(StrArg(args(0), who)), bs, ok)
        If ok Then
            out = x
        Else
            out = False
        End If
    Case P_RANDOM
        x = NumArg(args(0), who)
        If LIsExact(x) Then
            If x <= 0 Then LThrow2 "random: expected a positive integer, got", x
            out = LIntFromDouble(Int(RandDbl() * CDbl(x)))
        Else
            out = RandDbl() * CDbl(x)
        End If
    Case P_ADD1
        out = NumAdd(NumArg(args(0), who), 1&)
    Case P_SUB1
        out = NumSub(NumArg(args(0), who), 1&)
    Case Else
        LThrow "unknown primitive " & who
    End Select
End Sub

Private Function MakeValues2(ByRef a As Variant, ByRef b As Variant) As Variant
    Dim o As LispObj
    Set o = LNewObj(LK_VALUES)
    o.Push a
    o.Push b
    Set MakeValues2 = o
End Function

' ---------------------------------------------------------------------------
'  lists
' ---------------------------------------------------------------------------
Private Function CxrPath(ByVal nm As String) As String
    Select Case nm
    Case "first": CxrPath = "a"
    Case "second": CxrPath = "ad"
    Case "third": CxrPath = "add"
    Case "fourth": CxrPath = "addd"
    Case "fifth": CxrPath = "adddd"
    Case Else: CxrPath = Mid$(nm, 2, Len(nm) - 2)
    End Select
End Function

Private Function MemberGeneric(ByRef x As Variant, ByRef lst As Variant, ByVal mode As Long, ByRef cmpf As Variant, ByVal useCmp As Boolean) As Variant
    ' mode 0 eq, 1 eqv, 2 equal
    Dim y As Variant, p As LispObj, r As Variant, hit As Boolean
    LSetV y, lst
    Do While LIsPair(y)
        Set p = y
        If useCmp Then
            LApply2 cmpf, x, p.Car, r
            hit = LTruthy(r)
        ElseIf mode = 0 Then
            hit = LEqv(x, p.Car)
        ElseIf mode = 1 Then
            hit = LEqv(x, p.Car)
        Else
            hit = LEqual(x, p.Car)
        End If
        If hit Then
            Set MemberGeneric = p
            Exit Function
        End If
        LSetV y, p.Cdr
    Loop
    MemberGeneric = False
End Function

Private Function AssocGeneric(ByRef x As Variant, ByRef lst As Variant, ByVal mode As Long, ByRef cmpf As Variant, ByVal useCmp As Boolean) As Variant
    Dim y As Variant, p As LispObj, e As LispObj, r As Variant, hit As Boolean
    LSetV y, lst
    Do While LIsPair(y)
        Set p = y
        If LIsPair(p.Car) Then
            Set e = p.Car
            If useCmp Then
                LApply2 cmpf, x, e.Car, r
                hit = LTruthy(r)
            ElseIf mode = 2 Then
                hit = LEqual(x, e.Car)
            Else
                hit = LEqv(x, e.Car)
            End If
            If hit Then
                Set AssocGeneric = e
                Exit Function
            End If
        End If
        LSetV y, p.Cdr
    Loop
    AssocGeneric = False
End Function

' collect the cars of several lists at position k; False when one ends
Private Function NextCars(ByRef cur() As Variant, ByVal nl As Long, ByRef vals() As Variant) As Boolean
    Dim j As Long, p As LispObj
    For j = 0 To nl - 1
        If Not LIsPair(cur(j)) Then Exit Function
    Next
    For j = 0 To nl - 1
        Set p = cur(j)
        If IsObject(p.Car) Then
            Set vals(j) = p.Car
        Else
            vals(j) = p.Car
        End If
        If IsObject(p.Cdr) Then
            Set cur(j) = p.Cdr
        Else
            cur(j) = p.Cdr
        End If
    Next
    NextCars = True
End Function

Private Function KeyBefore(ByRef lessf As Variant, ByRef a As Variant, ByRef b As Variant) As Boolean
    Dim r As Variant
    LApply2 lessf, a, b, r
    KeyBefore = LTruthy(r)
End Function

Private Sub MergeSort(ByRef items() As Variant, ByRef tmp() As Variant, ByVal lo As Long, ByVal hi As Long, ByRef lessf As Variant)
    Dim md As Long, i As Long, j As Long, k As Long, t As Variant
    If hi - lo < 1 Then Exit Sub
    If hi - lo < 10 Then
        For i = lo + 1 To hi
            LSetV t, items(i)
            j = i - 1
            Do While j >= lo
                If Not KeyBefore(lessf, t, items(j)) Then Exit Do
                If IsObject(items(j)) Then
                    Set items(j + 1) = items(j)
                Else
                    items(j + 1) = items(j)
                End If
                j = j - 1
            Loop
            If IsObject(t) Then
                Set items(j + 1) = t
            Else
                items(j + 1) = t
            End If
        Next
        Exit Sub
    End If
    md = (lo + hi) \ 2
    MergeSort items, tmp, lo, md, lessf
    MergeSort items, tmp, md + 1, hi, lessf
    i = lo
    j = md + 1
    k = lo
    Do While i <= md And j <= hi
        If KeyBefore(lessf, items(j), items(i)) Then
            If IsObject(items(j)) Then
                Set tmp(k) = items(j)
            Else
                tmp(k) = items(j)
            End If
            j = j + 1
        Else
            If IsObject(items(i)) Then
                Set tmp(k) = items(i)
            Else
                tmp(k) = items(i)
            End If
            i = i + 1
        End If
        k = k + 1
    Loop
    Do While i <= md
        If IsObject(items(i)) Then
            Set tmp(k) = items(i)
        Else
            tmp(k) = items(i)
        End If
        i = i + 1
        k = k + 1
    Loop
    Do While j <= hi
        If IsObject(items(j)) Then
            Set tmp(k) = items(j)
        Else
            tmp(k) = items(j)
        End If
        j = j + 1
        k = k + 1
    Loop
    For k = lo To hi
        If IsObject(tmp(k)) Then
            Set items(k) = tmp(k)
        Else
            items(k) = tmp(k)
        End If
    Next
End Sub

Private Sub SortSeq(ByRef seq As Variant, ByRef lessf As Variant, ByVal who As String, ByRef out As Variant)
    Dim items() As Variant, tmp() As Variant, n As Long, i As Long, vec As LispObj, head As LispObj, tail As LispObj, r As LispObj
    ProcArg lessf, who
    If LIsKind(seq, LK_VECTOR) Then
        Set vec = seq
        n = vec.Count
        ReDim items(0 To IIf(n > 0, n - 1, 0))
        For i = 0 To n - 1
            vec.GetInto i, items(i)
        Next
    Else
        LListToArray seq, items, n, who
    End If
    If n > 1 Then
        ReDim tmp(0 To n - 1)
        MergeSort items, tmp, 0, n - 1, lessf
    End If
    If LIsKind(seq, LK_VECTOR) Then
        Set r = LNewObj(LK_VECTOR)
        For i = 0 To n - 1
            r.Push items(i)
        Next
        Set out = r
    Else
        For i = 0 To n - 1
            LB_Add head, tail, items(i)
        Next
        LB_Result head, out
    End If
End Sub

Private Sub PrimList(ByVal fo As LispObj, ByVal c As Long, ByRef args() As Variant, ByVal na As Long, ByRef out As Variant)
    Dim p As LispObj, x As Variant, y As Variant, i As Long, j As Long, n As Long, who As String, path As String
    Dim head As LispObj, tail As LispObj, r As Variant, cur() As Variant, vals() As Variant, nl As Long, acc As Variant
    Dim e As Variant, items() As Variant, ni As Long, st As Variant, stp As Variant, seen As LispObj, ks As String
    Dim yes As LispObj, yesTl As LispObj, no As LispObj, noTl As LispObj
    who = fo.SName
    Select Case c
    Case P_CONS
        Set out = LCons(args(0), args(1))
    Case P_CAR
        Set p = PairArg(args(0), who)
        LSetV out, p.Car
    Case P_CDR
        Set p = PairArg(args(0), who)
        LSetV out, p.Cdr
    Case P_SETCAR
        Set p = PairArg(args(0), who)
        p.SetCar args(1)
        Set out = lgVoid
    Case P_SETCDR
        Set p = PairArg(args(0), who)
        p.SetCdr args(1)
        Set out = lgVoid
    Case P_CXR
        path = CxrPath(who)
        LSetV x, args(0)
        For i = Len(path) To 1 Step -1
            If Not LIsPair(x) Then LThrow2 who & ": the list is too short:", args(0)
            Set p = x
            If Mid$(path, i, 1) = "a" Then
                LSetV x, p.Car
            Else
                LSetV x, p.Cdr
            End If
        Next
        LSetV out, x
    Case P_LIST
        LSetV out, LListFromArray(args, 0, na - 1)
    Case P_CONSSTAR
        LSetV r, args(na - 1)
        For i = na - 2 To 0 Step -1
            Set r = LCons(args(i), r)
        Next
        LSetV out, r
    Case P_LENGTH
        n = LLength(args(0))
        If n < 0 Then TypeErr who, "a proper list", args(0)
        out = n
    Case P_APPEND
        If na = 0 Then
            out = Empty
            Exit Sub
        End If
        LSetV r, args(na - 1)
        For i = na - 2 To 0 Step -1
            LSetV r, LAppend2(args(i), r)
        Next
        LSetV out, r
    Case P_REVERSE
        r = Empty
        LSetV x, args(0)
        Do While LIsPair(x)
            Set p = x
            Set r = LCons(p.Car, r)
            LSetV x, p.Cdr
        Loop
        If Not IsEmpty(x) Then TypeErr who, "a proper list", args(0)
        LSetV out, r
    Case P_LISTTAIL, P_LISTREF
        n = LongArg(args(1), who)
        LSetV x, args(0)
        For i = 1 To n
            If Not LIsPair(x) Then LThrow2 who & ": index out of range:", args(1)
            Set p = x
            LSetV x, p.Cdr
        Next
        If c = P_LISTTAIL Then
            LSetV out, x
        Else
            If Not LIsPair(x) Then LThrow2 who & ": index out of range:", args(1)
            Set p = x
            LSetV out, p.Car
        End If
    Case P_LISTCOPY
        LSetV x, args(0)
        Do While LIsPair(x)
            Set p = x
            LB_Add head, tail, p.Car
            LSetV x, p.Cdr
        Loop
        LB_Result head, out
        If Not (tail Is Nothing) Then tail.SetCdr x
    Case P_LASTPAIR, P_LAST
        Set p = PairArg(args(0), who)
        Do While LIsPair(p.Cdr)
            Set p = p.Cdr
        Loop
        If c = P_LAST Then
            LSetV out, p.Car
        Else
            Set out = p
        End If
    Case P_MEMQ
        LSetV out, MemberGeneric(args(0), args(1), 0, e, False)
    Case P_MEMV
        LSetV out, MemberGeneric(args(0), args(1), 1, e, False)
    Case P_MEMBER
        If na = 3 Then
            LSetV out, MemberGeneric(args(0), args(1), 2, args(2), True)
        Else
            LSetV out, MemberGeneric(args(0), args(1), 2, e, False)
        End If
    Case P_ASSQ, P_ASSV
        LSetV out, AssocGeneric(args(0), args(1), 1, e, False)
    Case P_ASSOC
        If na = 3 Then
            LSetV out, AssocGeneric(args(0), args(1), 2, args(2), True)
        Else
            LSetV out, AssocGeneric(args(0), args(1), 2, e, False)
        End If
    Case P_MAP, P_FOREACH, P_APPENDMAP, P_FILTERMAP
        ProcArg args(0), who
        nl = na - 1
        ReDim cur(0 To nl - 1)
        ReDim vals(0 To nl - 1)
        For j = 0 To nl - 1
            If IsObject(args(j + 1)) Then
                Set cur(j) = args(j + 1)
            Else
                cur(j) = args(j + 1)
            End If
        Next
        Do While NextCars(cur, nl, vals)
            LApply args(0), vals, nl, r
            Select Case c
            Case P_MAP
                LB_Add head, tail, r
            Case P_APPENDMAP
                LSetV x, r
                Do While LIsPair(x)
                    Set p = x
                    LB_Add head, tail, p.Car
                    LSetV x, p.Cdr
                Loop
            Case P_FILTERMAP
                If LTruthy(r) Then LB_Add head, tail, r
            End Select
        Loop
        If c = P_FOREACH Then
            Set out = lgVoid
        Else
            LB_Result head, out
        End If
    Case P_FILTER, P_REMOVE, P_PARTITION
        ProcArg args(0), who
        LSetV x, args(1)
        Do While LIsPair(x)
            Set p = x
            LApply1 args(0), p.Car, r
            If LTruthy(r) = (c <> P_REMOVE) Then
                LB_Add head, tail, p.Car
            ElseIf c = P_PARTITION Then
                LB_Add no, noTl, p.Car
            End If
            LSetV x, p.Cdr
        Loop
        If c = P_PARTITION Then
            LB_Result head, x
            LB_Result no, y
            LSetV out, MakeValues2(x, y)
        Else
            LB_Result head, out
        End If
    Case P_REDUCE
        ' (reduce f ridentity list): (f elem acc)
        ProcArg args(0), who
        LSetV x, args(2)
        If Not LIsPair(x) Then
            LSetV out, args(1)
            Exit Sub
        End If
        Set p = x
        LSetV acc, p.Car
        LSetV x, p.Cdr
        Do While LIsPair(x)
            Set p = x
            LApply2 args(0), p.Car, acc, r
            LSetV acc, r
            LSetV x, p.Cdr
        Loop
        LSetV out, acc
    Case P_FOLD, P_FOLDLEFT, P_FOLDRIGHT
        ProcArg args(0), who
        nl = na - 2
        ReDim cur(0 To nl - 1)
        ReDim vals(0 To nl)
        For j = 0 To nl - 1
            If c = P_FOLDRIGHT Then
                LSetV cur(j), ReverseList(args(j + 2))
            Else
                LSetV cur(j), args(j + 2)
            End If
        Next
        LSetV acc, args(1)
        Dim tv() As Variant
        ReDim tv(0 To nl - 1)
        Do While NextCars(cur, nl, tv)
            If c = P_FOLDLEFT Then
                LSetV vals(0), acc
                For j = 0 To nl - 1
                    LSetV vals(j + 1), tv(j)
                Next
            Else
                For j = 0 To nl - 1
                    LSetV vals(j), tv(j)
                Next
                LSetV vals(nl), acc
            End If
            LApply args(0), vals, nl + 1, r
            LSetV acc, r
        Loop
        LSetV out, acc
    Case P_DELETE
        LSetV x, args(1)
        Do While LIsPair(x)
            Set p = x
            If Not LEqual(args(0), p.Car) Then LB_Add head, tail, p.Car
            LSetV x, p.Cdr
        Loop
        LB_Result head, out
    Case P_LISTINDEX
        ProcArg args(0), who
        LSetV x, args(1)
        i = 0
        Do While LIsPair(x)
            Set p = x
            LApply1 args(0), p.Car, r
            If LTruthy(r) Then
                out = i
                Exit Sub
            End If
            i = i + 1
            LSetV x, p.Cdr
        Loop
        out = False
    Case P_IOTA
        n = LongArg(args(0), who)
        st = 0&
        stp = 1&
        If na >= 2 Then st = NumArg(args(1), who)
        If na >= 3 Then stp = NumArg(args(2), who)
        LSetV x, st
        For i = 1 To n
            LB_Add head, tail, x
            x = NumAdd(x, stp)
        Next
        LB_Result head, out
    Case LP_APPLY
        ProcArg args(0), who
        LSetV x, args(0)
        ReDim items(0 To na - 1)
        For i = 0 To na - 1
            If IsObject(args(i)) Then
                Set items(i) = args(i)
            Else
                items(i) = args(i)
            End If
        Next
        ni = na
        SpreadApplyArgs items, ni
        LApply x, items, ni, out
    Case P_SORT
        If LIsProc(args(0)) And Not LIsProc(args(1)) Then
            SortSeq args(1), args(0), who, out
        Else
            SortSeq args(0), args(1), who, out
        End If
    Case P_LISTSORT
        SortSeq args(1), args(0), who, out
    Case P_ANY, P_EVERY
        ProcArg args(0), who
        nl = na - 1
        ReDim cur(0 To nl - 1)
        ReDim vals(0 To nl - 1)
        For j = 0 To nl - 1
            If IsObject(args(j + 1)) Then
                Set cur(j) = args(j + 1)
            Else
                cur(j) = args(j + 1)
            End If
        Next
        If c = P_ANY Then
            out = False
        Else
            out = True
        End If
        Do While NextCars(cur, nl, vals)
            LApply args(0), vals, nl, r
            If c = P_ANY Then
                If LTruthy(r) Then
                    LSetV out, r
                    Exit Sub
                End If
            Else
                If Not LTruthy(r) Then
                    out = False
                    Exit Sub
                End If
                LSetV out, r
            End If
        Loop
    Case P_COUNT
        ProcArg args(0), who
        n = 0
        LSetV x, args(1)
        Do While LIsPair(x)
            Set p = x
            LApply1 args(0), p.Car, r
            If LTruthy(r) Then n = n + 1
            LSetV x, p.Cdr
        Loop
        out = n
    Case P_DELDUP
        Set seen = LNewObj(LK_HASH)
        LSetV x, args(0)
        Do While LIsPair(x)
            Set p = x
            If na = 2 Then
                If Not LTruthy(MemberGeneric(p.Car, ListFromBuilder(head), 2, args(1), True)) Then LB_Add head, tail, p.Car
            Else
                ks = LKeyStr(p.Car)
                If seen.FindH(ks, LHashStr(ks)) < 0 Then
                    seen.SetH ks, LHashStr(ks), e, e
                    LB_Add head, tail, p.Car
                End If
            End If
            LSetV x, p.Cdr
        Loop
        LB_Result head, out
    Case P_FIND
        ProcArg args(0), who
        LSetV x, args(1)
        Do While LIsPair(x)
            Set p = x
            LApply1 args(0), p.Car, r
            If LTruthy(r) Then
                LSetV out, p.Car
                Exit Sub
            End If
            LSetV x, p.Cdr
        Loop
        out = False
    Case P_LISTP
        out = (LLength(args(0)) >= 0)
    Case P_PAIRP
        out = LIsPair(args(0))
    Case P_NULLP
        out = IsEmpty(args(0))
    Case P_MAKELIST
        n = LongArg(args(0), who)
        If na = 2 Then
            LSetV x, args(1)
        Else
            Set x = lgVoid
        End If
        r = Empty
        For i = 1 To n
            Set r = LCons(x, r)
        Next
        LSetV out, r
    Case P_LISTSET
        n = LongArg(args(1), who)
        LSetV x, args(0)
        For i = 1 To n
            If Not LIsPair(x) Then LThrow2 who & ": index out of range:", args(1)
            Set p = x
            LSetV x, p.Cdr
        Next
        Set p = PairArg(x, who)
        p.SetCar args(2)
        Set out = lgVoid
    Case Else
        LThrow "unknown primitive " & who
    End Select
End Sub

Private Function ListFromBuilder(ByVal head As LispObj) As Variant
    If head Is Nothing Then
        ListFromBuilder = Empty
    Else
        Set ListFromBuilder = head
    End If
End Function

Private Function ReverseList(ByRef lst As Variant) As Variant
    Dim r As Variant, x As Variant, p As LispObj
    LSetV x, lst
    Do While LIsPair(x)
        Set p = x
        Set r = LCons(p.Car, r)
        LSetV x, p.Cdr
    Loop
    If IsObject(r) Then Set ReverseList = r Else ReverseList = r
End Function

' ---------------------------------------------------------------------------
'  strings, characters, symbols
' ---------------------------------------------------------------------------
Private Function IsWsChar(ByVal c As String) As Boolean
    Select Case LUCode(c)
    Case 9 To 13, 32, 160, &H3000&
        IsWsChar = True
    End Select
End Function

Private Function TrimWs(ByVal s As String, ByVal mode As Long) As String
    Dim i As Long, j As Long
    i = 1
    j = Len(s)
    If (mode And 1) <> 0 Then
        Do While i <= j
            If Not IsWsChar(Mid$(s, i, 1)) Then Exit Do
            i = i + 1
        Loop
    End If
    If (mode And 2) <> 0 Then
        Do While j >= i
            If Not IsWsChar(Mid$(s, j, 1)) Then Exit Do
            j = j - 1
        Loop
    End If
    If j < i Then
        TrimWs = vbNullString
    Else
        TrimWs = Mid$(s, i, j - i + 1)
    End If
End Function

Private Function IsAlphaCh(ByVal c As String) As Boolean
    Dim k As Long
    k = LUCode(c)
    If (k >= 65 And k <= 90) Or (k >= 97 And k <= 122) Then
        IsAlphaCh = True
    ElseIf k >= 128 Then
        If UCase$(c) <> LCase$(c) Then
            IsAlphaCh = True
        ElseIf (k >= &H4E00& And k <= &H9FFF&) Or (k >= &H3400& And k <= &H4DBF&) Or (k >= &H3040& And k <= &H30FF&) Or (k >= &HAC00& And k <= &HD7AF&) Then
            IsAlphaCh = True
        End If
    End If
End Function

Private Function StrCmpSeq(ByRef args() As Variant, ByVal na As Long, ByVal who As String, ByVal mode As Long, ByVal ci As Boolean) As Boolean
    ' mode: 0 =, 1 <, 2 >, 3 <=, 4 >=
    Dim i As Long, a As String, b As String, cmp As Long
    a = StrArg(args(0), who)
    For i = 1 To na - 1
        b = StrArg(args(i), who)
        If ci Then
            cmp = StrComp(LCase$(a), LCase$(b), vbBinaryCompare)
        Else
            cmp = StrComp(a, b, vbBinaryCompare)
        End If
        Select Case mode
        Case 0: If cmp <> 0 Then Exit Function
        Case 1: If cmp >= 0 Then Exit Function
        Case 2: If cmp <= 0 Then Exit Function
        Case 3: If cmp > 0 Then Exit Function
        Case 4: If cmp < 0 Then Exit Function
        End Select
        a = b
    Next
    StrCmpSeq = True
End Function

Private Function CharCmpSeq(ByRef args() As Variant, ByVal na As Long, ByVal who As String, ByVal mode As Long) As Boolean
    Dim i As Long, a As Long, b As Long, ok As Boolean
    a = LUCode(CharArg(args(0), who))
    For i = 1 To na - 1
        b = LUCode(CharArg(args(i), who))
        Select Case mode
        Case 0: ok = (a = b)
        Case 1: ok = (a < b)
        Case 2: ok = (a > b)
        Case 3: ok = (a <= b)
        Case Else: ok = (a >= b)
        End Select
        If Not ok Then Exit Function
        a = b
    Next
    CharCmpSeq = True
End Function

Private Sub SubRange(ByVal n As Long, ByRef args() As Variant, ByVal na As Long, ByVal first As Long, ByVal who As String, ByRef st As Long, ByRef en As Long)
    st = 0
    en = n
    If na > first Then st = LongArg(args(first), who)
    If na > first + 1 Then en = LongArg(args(first + 1), who)
    If st < 0 Or en > n Or st > en Then LThrow who & ": index out of range (" & st & ", " & en & ") for length " & n
End Sub

Private Sub PrimStr(ByVal fo As LispObj, ByVal c As Long, ByRef args() As Variant, ByVal na As Long, ByRef out As Variant)
    Dim s As String, t As String, i As Long, n As Long, who As String, st As Long, en As Long, o As LispObj
    Dim x As Variant, p As LispObj, head As LispObj, tail As LispObj, parts() As String, r As Variant, k As Long, ch As String
    who = fo.SName
    Select Case c
    Case P_SYMBOLP
        out = LIsSym(args(0))
    Case P_STRINGP
        out = (VarType(args(0)) = vbString)
    Case P_CHARP
        out = LIsKind(args(0), LK_CHAR)
    Case P_SYM2STR
        Set o = ObjArg(args(0), LK_SYM, who, "a symbol")
        out = o.SName
    Case P_STR2SYM
        Set out = LIntern(StrArg(args(0), who))
    Case P_GENSYM
        lgGensym = lgGensym + 1
        Set o = LNewObj(LK_SYM)
        t = "g"
        If na = 1 Then t = LToStr(args(0), False)
        o.SName = t & lgGensym
        o.H = LHashStr(o.SName)
        Set out = o
    Case P_STRLEN
        out = CLng(Len(StrArg(args(0), who)))
    Case P_STRREF
        s = StrArg(args(0), who)
        k = LongArg(args(1), who)
        If k < 0 Or k >= Len(s) Then LThrow2 who & ": index out of range:", args(1)
        Set out = LMakeChar(Mid$(s, k + 1, 1))
    Case P_SUBSTRING, P_STRCOPY
        s = StrArg(args(0), who)
        SubRange Len(s), args, na, 1, who, st, en
        out = Mid$(s, st + 1, en - st)
    Case P_STRAPPEND
        If na = 0 Then
            out = vbNullString
        Else
            ReDim parts(0 To na - 1)
            For i = 0 To na - 1
                parts(i) = StrArg(args(i), who)
            Next
            out = Join(parts, vbNullString)
        End If
    Case P_STREQ
        out = StrCmpSeq(args, na, who, 0, False)
    Case P_STRLT
        out = StrCmpSeq(args, na, who, 1, False)
    Case P_STRGT
        out = StrCmpSeq(args, na, who, 2, False)
    Case P_STRLE
        out = StrCmpSeq(args, na, who, 3, False)
    Case P_STRGE
        out = StrCmpSeq(args, na, who, 4, False)
    Case P_STRCIEQ
        out = StrCmpSeq(args, na, who, 0, True)
    Case P_STRUP
        out = UCase$(StrArg(args(0), who))
    Case P_STRDOWN
        out = LCase$(StrArg(args(0), who))
    Case P_STR2LIST
        s = StrArg(args(0), who)
        SubRange Len(s), args, na, 1, who, st, en
        For i = st + 1 To en
            LB_Add head, tail, ToV(LMakeChar(Mid$(s, i, 1)))
        Next
        LB_Result head, out
    Case P_LIST2STR
        n = LLength(args(0))
        If n < 0 Then TypeErr who, "a list of characters", args(0)
        If n = 0 Then
            out = vbNullString
        Else
            ReDim parts(0 To n - 1)
            LSetV x, args(0)
            i = 0
            Do While LIsPair(x)
                Set p = x
                parts(i) = CharArg(p.Car, who)
                i = i + 1
                LSetV x, p.Cdr
            Loop
            out = Join(parts, vbNullString)
        End If
    Case P_STRING
        If na = 0 Then
            out = vbNullString
        Else
            ReDim parts(0 To na - 1)
            For i = 0 To na - 1
                If LIsKind(args(i), LK_CHAR) Then
                    parts(i) = CharArg(args(i), who)
                Else
                    parts(i) = LToStr(args(i), False)
                End If
            Next
            out = Join(parts, vbNullString)
        End If
    Case P_MAKESTR
        n = LongArg(args(0), who)
        ch = " "
        If na = 2 Then ch = CharArg(args(1), who)
        If n <= 0 Then
            out = vbNullString
        Else
            out = String$(n, ch)
        End If
    Case P_STRNULLP
        out = (Len(StrArg(args(0), who)) = 0)
    Case P_STRINDEX
        s = StrArg(args(0), who)
        out = False
        For i = 1 To Len(s)
            If LIsKind(args(1), LK_CHAR) Then
                If Mid$(s, i, 1) = CharArg(args(1), who) Then
                    out = CLng(i - 1)
                    Exit Sub
                End If
            Else
                LApply1 args(1), ToV(LMakeChar(Mid$(s, i, 1))), r
                If LTruthy(r) Then
                    out = CLng(i - 1)
                    Exit Sub
                End If
            End If
        Next
    Case P_STRCONTAINS
        s = StrArg(args(0), who)
        t = StrArg(args(1), who)
        k = InStr(1, s, t, vbBinaryCompare)
        If Len(t) = 0 Then
            out = 0&
        ElseIf k = 0 Then
            out = False
        Else
            out = CLng(k - 1)
        End If
    Case P_STRCONTAINSP
        s = StrArg(args(0), who)
        t = StrArg(args(1), who)
        out = (Len(t) = 0 Or InStr(1, s, t, vbBinaryCompare) > 0)
    Case P_STRSPLIT
        s = StrArg(args(0), who)
        If na = 1 Then
            ' split on runs of whitespace (Racket)
            i = 1
            Do
                Do While i <= Len(s)
                    If Not IsWsChar(Mid$(s, i, 1)) Then Exit Do
                    i = i + 1
                Loop
                If i > Len(s) Then Exit Do
                k = i
                Do While k <= Len(s)
                    If IsWsChar(Mid$(s, k, 1)) Then Exit Do
                    k = k + 1
                Loop
                x = Mid$(s, i, k - i)
                LB_Add head, tail, x
                i = k
            Loop
        Else
            If LIsKind(args(1), LK_CHAR) Then
                t = CharArg(args(1), who)
            Else
                t = StrArg(args(1), who)
            End If
            If Len(t) = 0 Then LThrow who & ": empty separator"
            i = 1
            Do
                k = InStr(i, s, t, vbBinaryCompare)
                If k = 0 Then Exit Do
                x = Mid$(s, i, k - i)
                LB_Add head, tail, x
                i = k + Len(t)
            Loop
            x = Mid$(s, i)
            LB_Add head, tail, x
        End If
        LB_Result head, out
    Case P_STRJOIN
        t = " "
        If na = 2 Then t = StrArg(args(1), who)
        n = LLength(args(0))
        If n < 0 Then TypeErr who, "a list of strings", args(0)
        If n = 0 Then
            out = vbNullString
        Else
            ReDim parts(0 To n - 1)
            LSetV x, args(0)
            i = 0
            Do While LIsPair(x)
                Set p = x
                parts(i) = StrArg(p.Car, who)
                i = i + 1
                LSetV x, p.Cdr
            Loop
            out = Join(parts, t)
        End If
    Case P_STRTRIM
        out = TrimWs(StrArg(args(0), who), 3)
    Case P_STRTRIML
        out = TrimWs(StrArg(args(0), who), 1)
    Case P_STRTRIMR
        out = TrimWs(StrArg(args(0), who), 2)
    Case P_STRREV
        s = StrArg(args(0), who)
        n = Len(s)
        If n = 0 Then
            out = vbNullString
        Else
            ReDim parts(0 To n - 1)
            For i = 1 To n
                parts(n - i) = Mid$(s, i, 1)
            Next
            out = Join(parts, vbNullString)
        End If
    Case P_STRPREFIXP
        t = StrArg(args(0), who)
        s = StrArg(args(1), who)
        out = (Left$(s, Len(t)) = t And Len(t) <= Len(s))
    Case P_STRSUFFIXP
        t = StrArg(args(0), who)
        s = StrArg(args(1), who)
        out = (Right$(s, Len(t)) = t And Len(t) <= Len(s))
    Case P_STRFOREACH, P_STRMAP
        ProcArg args(0), who
        s = StrArg(args(1), who)
        If Len(s) > 0 Then
            ReDim parts(0 To Len(s) - 1)
        End If
        For i = 1 To Len(s)
            LApply1 args(0), ToV(LMakeChar(Mid$(s, i, 1))), r
            If c = P_STRMAP Then parts(i - 1) = CharArg(r, who)
        Next
        If c = P_STRMAP Then
            If Len(s) = 0 Then
                out = vbNullString
            Else
                out = Join(parts, vbNullString)
            End If
        Else
            Set out = lgVoid
        End If
    Case P_STRPAD, P_STRPADR
        s = StrArg(args(0), who)
        n = LongArg(args(1), who)
        ch = " "
        If na = 3 Then ch = CharArg(args(2), who)
        If c = P_STRPAD Then
            If Len(s) >= n Then
                out = Right$(s, n)
            Else
                out = String$(n - Len(s), ch) & s
            End If
        Else
            If Len(s) >= n Then
                out = Left$(s, n)
            Else
                out = s & String$(n - Len(s), ch)
            End If
        End If
    Case P_STRREPLACE
        s = StrArg(args(0), who)
        t = StrArg(args(1), who)
        If Len(t) = 0 Then LThrow who & ": empty pattern"
        out = Replace(s, t, StrArg(args(2), who), 1, -1, vbBinaryCompare)
    Case P_CHAR2INT
        out = LUCode(CharArg(args(0), who))
    Case P_INT2CHAR
        k = LongArg(args(0), who)
        If k < 0 Or k > 65535 Then LThrow2 who & ": code point out of range:", args(0)
        Set out = LMakeChar(ChrW(k))
    Case P_CHARUP
        Set out = LMakeChar(UCase$(CharArg(args(0), who)))
    Case P_CHARDOWN
        Set out = LMakeChar(LCase$(CharArg(args(0), who)))
    Case P_CHARALPHAP
        out = IsAlphaCh(CharArg(args(0), who))
    Case P_CHARNUMP
        ch = CharArg(args(0), who)
        out = (ch >= "0" And ch <= "9")
    Case P_CHARWSP
        out = IsWsChar(CharArg(args(0), who))
    Case P_CHARUPPERP
        ch = CharArg(args(0), who)
        out = (UCase$(ch) = ch And LCase$(ch) <> ch)
    Case P_CHARLOWERP
        ch = CharArg(args(0), who)
        out = (LCase$(ch) = ch And UCase$(ch) <> ch)
    Case P_DIGITVAL
        ch = CharArg(args(0), who)
        If ch >= "0" And ch <= "9" Then out = CLng(Asc(ch) - 48) Else out = False
    Case P_CHAREQ
        out = CharCmpSeq(args, na, who, 0)
    Case P_CHARLT
        out = CharCmpSeq(args, na, who, 1)
    Case P_CHARGT
        out = CharCmpSeq(args, na, who, 2)
    Case P_CHARLE
        out = CharCmpSeq(args, na, who, 3)
    Case P_CHARGE
        out = CharCmpSeq(args, na, who, 4)
    Case Else
        LThrow "unknown primitive " & who
    End Select
End Sub

' ---------------------------------------------------------------------------
'  vectors and hash tables
' ---------------------------------------------------------------------------
Private Function VecIndex(ByVal vec As LispObj, ByRef k As Variant, ByVal who As String) As Long
    Dim i As Long
    i = LongArg(k, who)
    If i < 0 Or i >= vec.Count Then LThrow2 who & ": index out of range:", k
    VecIndex = i
End Function

Private Function HashFind(ByVal h As LispObj, ByRef key As Variant) As Long
    Dim ks As String
    ks = LKeyStr(key)
    HashFind = h.FindH(ks, LHashStr(ks))
End Function

Private Sub HashPut(ByVal h As LispObj, ByRef key As Variant, ByRef v As Variant)
    Dim ks As String
    ks = LKeyStr(key)
    h.SetH ks, LHashStr(ks), key, v
End Sub

Private Sub PrimVec(ByVal fo As LispObj, ByVal c As Long, ByRef args() As Variant, ByVal na As Long, ByRef out As Variant)
    Dim v As LispObj, r As LispObj, i As Long, n As Long, x As Variant, p As LispObj, who As String, st As Long, en As Long
    Dim h As LispObj, e As Long, head As LispObj, tail As LispObj, k As Variant, y As Variant, fx As Variant
    Dim vals() As Variant, j As Long, nl As Long, vecs() As LispObj, ok As Boolean, ev As Variant
    who = fo.SName
    Select Case c
    Case P_VECTORP
        out = LIsKind(args(0), LK_VECTOR)
    Case P_MAKEVECTOR
        n = LongArg(args(0), who)
        If na = 2 Then
            LSetV x, args(1)
        Else
            x = 0&
        End If
        Set r = LNewObj(LK_VECTOR)
        r.Reserve n
        For i = 1 To n
            r.Push x
        Next
        Set out = r
    Case P_VECTOR
        Set r = LNewObj(LK_VECTOR)
        For i = 0 To na - 1
            r.Push args(i)
        Next
        Set out = r
    Case P_VECREF
        Set v = ObjArg(args(0), LK_VECTOR, who, "a vector")
        v.GetInto VecIndex(v, args(1), who), out
    Case P_VECSET
        Set v = ObjArg(args(0), LK_VECTOR, who, "a vector")
        v.SetI VecIndex(v, args(1), who), args(2)
        Set out = lgVoid
    Case P_VECLEN
        Set v = ObjArg(args(0), LK_VECTOR, who, "a vector")
        out = v.Count
    Case P_VEC2LIST
        Set v = ObjArg(args(0), LK_VECTOR, who, "a vector")
        SubRange v.Count, args, na, 1, who, st, en
        For i = st To en - 1
            v.GetInto i, x
            LB_Add head, tail, x
        Next
        LB_Result head, out
    Case P_LIST2VEC
        If LLength(args(0)) < 0 Then TypeErr who, "a proper list", args(0)
        Set r = LNewObj(LK_VECTOR)
        LSetV x, args(0)
        Do While LIsPair(x)
            Set p = x
            r.Push p.Car
            LSetV x, p.Cdr
        Loop
        Set out = r
    Case P_VECFILL
        Set v = ObjArg(args(0), LK_VECTOR, who, "a vector")
        For i = 0 To v.Count - 1
            v.SetI i, args(1)
        Next
        Set out = lgVoid
    Case P_VECCOPY
        Set v = ObjArg(args(0), LK_VECTOR, who, "a vector")
        SubRange v.Count, args, na, 1, who, st, en
        Set r = LNewObj(LK_VECTOR)
        For i = st To en - 1
            v.GetInto i, x
            r.Push x
        Next
        Set out = r
    Case P_VECMAP, P_VECFOREACH
        ProcArg args(0), who
        nl = na - 1
        ReDim vecs(0 To nl - 1)
        ReDim vals(0 To nl - 1)
        n = -1
        For j = 0 To nl - 1
            Set v = ObjArg(args(j + 1), LK_VECTOR, who, "a vector")
            Set vecs(j) = v
            If n < 0 Or v.Count < n Then n = v.Count
        Next
        Set r = LNewObj(LK_VECTOR)
        For i = 0 To n - 1
            For j = 0 To nl - 1
                Set v = vecs(j)
                v.GetInto i, x
                If IsObject(x) Then
                    Set vals(j) = x
                Else
                    vals(j) = x
                End If
            Next
            LApply args(0), vals, nl, y
            If c = P_VECMAP Then r.Push y
        Next
        If c = P_VECMAP Then
            Set out = r
        Else
            Set out = lgVoid
        End If
    Case P_MAKEHASH
        Set out = LNewObj(LK_HASH)
    Case P_HASHP
        out = LIsKind(args(0), LK_HASH)
    Case P_HASHSET
        Set h = ObjArg(args(0), LK_HASH, who, "a hash table")
        HashPut h, args(1), args(2)
        Set out = lgVoid
    Case P_HASHREF, P_HASHREFDEF, P_HASHGUILEREF
        Set h = ObjArg(args(0), LK_HASH, who, "a hash table")
        e = HashFind(h, args(1))
        If e >= 0 Then
            h.ValInto e, out
        ElseIf c = P_HASHREFDEF Then
            LSetV out, args(2)
        ElseIf c = P_HASHGUILEREF Then
            If na = 3 Then
                LSetV out, args(2)
            Else
                out = False
            End If
        ElseIf na = 3 Then
            LApply0 args(2), out
        Else
            LThrow2 who & ": key not found:", args(1)
        End If
    Case P_HASHDEL
        Set h = ObjArg(args(0), LK_HASH, who, "a hash table")
        e = HashFind(h, args(1))
        If e >= 0 Then h.DelAt e
        Set out = lgVoid
    Case P_HASHHAS
        Set h = ObjArg(args(0), LK_HASH, who, "a hash table")
        out = (HashFind(h, args(1)) >= 0)
    Case P_HASHCOUNT
        Set h = ObjArg(args(0), LK_HASH, who, "a hash table")
        out = h.DCount
    Case P_HASHKEYS, P_HASHVALUES, P_HASH2ALIST
        Set h = ObjArg(args(0), LK_HASH, who, "a hash table")
        For i = 0 To h.EntryCount - 1
            If h.LiveAt(i) Then
                h.KeyInto i, k
                h.ValInto i, y
                Select Case c
                Case P_HASHKEYS: LB_Add head, tail, k
                Case P_HASHVALUES: LB_Add head, tail, y
                Case Else: LB_Add head, tail, ToV(LCons(k, y))
                End Select
            End If
        Next
        LB_Result head, out
    Case P_HASHUPDATE, P_HASHUPDATEDEF
        Set h = ObjArg(args(0), LK_HASH, who, "a hash table")
        ProcArg args(2), who
        e = HashFind(h, args(1))
        If e >= 0 Then
            h.ValInto e, x
        ElseIf c = P_HASHUPDATEDEF Then
            LSetV x, args(3)
        ElseIf na = 4 Then
            LApply0 args(3), x
        Else
            LThrow2 who & ": key not found:", args(1)
        End If
        LApply1 args(2), x, y
        HashPut h, args(1), y
        Set out = lgVoid
    Case P_HASHWALK, P_HASHFOREACH
        If c = P_HASHWALK Then
            Set h = ObjArg(args(0), LK_HASH, who, "a hash table")
            LSetV fx, args(1)
        Else
            Set h = ObjArg(args(1), LK_HASH, who, "a hash table")
            LSetV fx, args(0)
        End If
        ProcArg fx, who
        For i = 0 To h.EntryCount - 1
            If h.LiveAt(i) Then
                h.KeyInto i, k
                h.ValInto i, y
                LApply2 fx, k, y, x
            End If
        Next
        Set out = lgVoid
    Case P_HASHCLEAR
        Set h = ObjArg(args(0), LK_HASH, who, "a hash table")
        h.DClear
        Set out = lgVoid
    Case P_HASHCOPY
        Set h = ObjArg(args(0), LK_HASH, who, "a hash table")
        Set r = LNewObj(LK_HASH)
        For i = 0 To h.EntryCount - 1
            If h.LiveAt(i) Then
                h.KeyInto i, k
                h.ValInto i, y
                r.SetH h.KSAt(i), h.HAt(i), k, y
            End If
        Next
        Set out = r
    Case Else
        LThrow "unknown primitive " & who
    End Select
End Sub

' ---------------------------------------------------------------------------
'  control, output, errors
' ---------------------------------------------------------------------------
Private Function FormatStr(ByVal fmt As String, ByRef args() As Variant, ByVal first As Long, ByVal na As Long) As String
    Dim i As Long, n As Long, c As String, parts() As String, np As Long, ai As Long, p As Long
    n = Len(fmt)
    ReDim parts(0 To 15)
    ai = first
    i = 1
    Do While i <= n
        p = InStr(i, fmt, "~")
        If p = 0 Then
            AddPart parts, np, Mid$(fmt, i)
            Exit Do
        End If
        AddPart parts, np, Mid$(fmt, i, p - i)
        c = LCase$(Mid$(fmt, p + 1, 1))
        i = p + 2
        Select Case c
        Case "a", "s", "d", "b", "o", "x", "v"
            If ai >= na Then LThrow "format: too few arguments for " & LToStr(fmt, True)
            Select Case c
            Case "a", "d"
                AddPart parts, np, LToStr(args(ai), False)
            Case "s", "v"
                AddPart parts, np, LToStr(args(ai), True)
            Case "b"
                AddPart parts, np, IntToRadix(IntArg(args(ai), "format"), 2)
            Case "o"
                AddPart parts, np, IntToRadix(IntArg(args(ai), "format"), 8)
            Case "x"
                AddPart parts, np, IntToRadix(IntArg(args(ai), "format"), 16)
            End Select
            ai = ai + 1
        Case "%", "n"
            AddPart parts, np, vbLf
        Case "~"
            AddPart parts, np, "~"
        Case Else
            LThrow "format: unknown directive ~" & c
        End Select
    Loop
    If np = 0 Then Exit Function
    ReDim Preserve parts(0 To np - 1)
    FormatStr = Join(parts, vbNullString)
End Function

Private Sub AddPart(ByRef parts() As String, ByRef np As Long, ByVal s As String)
    If np > UBound(parts) Then
        ReDim Preserve parts(0 To 2 * np + 1)
    End If
    parts(np) = s
    np = np + 1
End Sub

Private Sub PrimCtl(ByVal fo As LispObj, ByVal c As Long, ByRef args() As Variant, ByVal na As Long, ByRef out As Variant)
    Dim o As LispObj, x As Variant, y As Variant, i As Long, who As String, s As String, msg As String, irr As Variant
    Dim vals() As Variant, nv As Long, p As LispObj, f As Variant
    who = fo.SName
    Select Case c
    Case P_PROCP
        out = LIsProc(args(0))
    Case P_BOOLP
        out = (VarType(args(0)) = vbBoolean)
    Case P_NOT
        out = Not LTruthy(args(0))
    Case P_EQP, P_EQVP
        out = LEqv(args(0), args(1))
    Case P_EQUALP
        out = LEqual(args(0), args(1))
    Case P_SYMEQ, P_BOOLEQ
        out = True
        For i = 1 To na - 1
            If Not LEqv(args(0), args(i)) Then out = False
        Next
    Case P_DISPLAY
        LOutWrite LToStr(args(0), False)
        Set out = lgVoid
    Case P_WRITE
        LOutWrite LToStr(args(0), True)
        Set out = lgVoid
    Case P_NEWLINE
        LOutWrite vbLf
        Set out = lgVoid
    Case P_WRITESTR
        LOutWrite StrArg(args(0), who)
        Set out = lgVoid
    Case P_WRITECHAR
        LOutWrite CharArg(args(0), who)
        Set out = lgVoid
    Case P_DISPLAYLN, P_WRITELINE
        LOutWrite LToStr(args(0), False) & vbLf
        Set out = lgVoid
    Case P_FORMAT
        If VarType(args(0)) = vbString Then
            out = FormatStr(args(0), args, 1, na)
        Else
            If na < 2 Then LThrow "format: missing format string"
            s = FormatStr(StrArg(args(1), who), args, 2, na)
            If VarType(args(0)) = vbBoolean Then
                If args(0) Then
                    LOutWrite s
                    Set out = lgVoid
                Else
                    out = s
                End If
            Else
                LOutWrite s
                Set out = lgVoid
            End If
        End If
    Case P_ERROR
        If na = 0 Then
            msg = "error"
            irr = Empty
        ElseIf VarType(args(0)) = vbString Then
            msg = args(0)
            LSetV irr, LListFromArray(args, 1, na - 1)
        ElseIf LIsSym(args(0)) And na >= 2 Then
            If VarType(args(1)) = vbString Then
                msg = LToStr(args(0), False) & ": " & args(1)
                LSetV irr, LListFromArray(args, 2, na - 1)
            Else
                msg = LToStr(args(0), False)
                LSetV irr, LListFromArray(args, 1, na - 1)
            End If
        Else
            msg = LToStr(args(0), False)
            LSetV irr, LListFromArray(args, 1, na - 1)
        End If
        LRaise LMakeError(msg, irr)
    Case P_RAISE
        LRaise args(0)
    Case P_WITHHANDLER
        ProcArg args(0), who
        ProcArg args(1), who
        LWithHandler args(0), args(1), out
    Case P_ERROBJP
        out = LIsKind(args(0), LK_ERROR)
    Case P_ERRMSG
        If LIsKind(args(0), LK_ERROR) Then
            Set o = args(0)
            out = o.SName
        Else
            out = LToStr(args(0), False)
        End If
    Case P_ERRIRR
        If LIsKind(args(0), LK_ERROR) Then
            Set o = args(0)
            LSetV out, o.Car
        Else
            out = Empty
        End If
    Case P_REPORTSTR
        out = LCondReport(args(0))
    Case P_EXIT
        Set o = LMakeError("exit", irr)
        o.Flags = 99
        LRaise o
    Case P_VOID
        Set out = lgVoid
    Case P_FORCE
        LForce args(0), out
    Case P_MAKEPROMISE
        If LIsKind(args(0), LK_PROMISE) Then
            LSetV out, args(0)
        Else
            Set o = LNewObj(LK_PROMISE)
            o.Flags = 1
            o.SetCdr args(0)
            Set out = o
        End If
    Case P_PROMISEP
        out = LIsKind(args(0), LK_PROMISE)
    Case P_CALLCC
        ProcArg args(0), who
        LCallCC args(0), out
    Case P_VALUES
        If na = 0 Then
            Set out = LNewObj(LK_VALUES)
        Else
            LSetV out, MakeValues(args, na)
        End If
    Case P_CALLWV
        ProcArg args(0), who
        ProcArg args(1), who
        LApply0 args(0), x
        If LIsKind(x, LK_VALUES) Then
            Set o = x
            nv = o.Count
            ReDim vals(0 To IIf(nv > 0, nv - 1, 0))
            For i = 0 To nv - 1
                o.GetInto i, y
                If IsObject(y) Then
                    Set vals(i) = y
                Else
                    vals(i) = y
                End If
            Next
        Else
            nv = 1
            ReDim vals(0 To 0)
            If IsObject(x) Then
                Set vals(0) = x
            Else
                vals(0) = x
            End If
        End If
        LApply args(1), vals, nv, out
    Case P_DYNWIND
        LDynamicWind args(0), args(1), args(2), out
    Case P_EVAL
        If na = 2 And LIsKind(args(1), LK_ENV) Then
            Set o = args(1)
            LEval args(0), o, out
        Else
            LEval args(0), lgGlobal, out
        End If
    Case P_INTERENV
        Set out = lgGlobal
    Case P_RUNTIME
        out = CDbl(Int(Timer * 1000#))
    Case P_CURSECONDS
        out = LIntFromDouble(Int((CDbl(Date) - 25569#) * 86400# + CDbl(Timer)))
    Case P_STREAMCAR
        Set p = PairArg(args(0), who)
        LSetV out, p.Car
    Case P_STREAMCDR
        Set p = PairArg(args(0), who)
        LForce p.Cdr, out
    Case P_STREAMPAIRP
        out = False
        If LIsPair(args(0)) Then
            Set p = args(0)
            out = LIsKind(p.Cdr, LK_PROMISE)
        End If
    Case P_READLINE
        out = LispHostInput("read-line")
    Case P_READ
        LSetV x, LReadAll(LispHostInput("read"))
        If LIsPair(x) Then
            Set p = x
            LSetV out, p.Car
        Else
            Set out = lgEof
        End If
    Case P_IDENTITY
        LSetV out, args(0)
    Case P_PROCARITY
        If LIsKind(args(0), LK_LAMBDA) Then
            Set o = args(0)
            If o.Ref Is Nothing Then
                out = o.I1
            Else
                Set out = LCons(o.I1, False)
            End If
        ElseIf LIsKind(args(0), LK_PRIM) Then
            Set o = args(0)
            If o.I2 = o.I1 Then
                out = o.I1
            ElseIf o.I2 < 0 Then
                Set out = LCons(o.I1, False)
            Else
                Set out = LCons(o.I1, o.I2)
            End If
        Else
            TypeErr who, "a procedure", args(0)
        End If
    Case Else
        LThrow "unknown primitive " & who
    End Select
End Sub
