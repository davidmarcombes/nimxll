# From xlcall.h
# Initially generated by c2nim

# Removed _
# addded _2 to duplicate names

##
## * Built-in Excel functions and command equivalents
##
##  Excel function numbers

const
  xlfCount* = 0
  xlfIsna* = 2
  xlfIserror* = 3
  xlfSum* = 4
  xlfAverage* = 5
  xlfMin* = 6
  xlfMax* = 7
  xlfRow* = 8
  xlfColumn* = 9
  xlfNa* = 10
  xlfNpv* = 11
  xlfStdev* = 12
  xlfDollar* = 13
  xlfFixed* = 14
  xlfSin* = 15
  xlfCos* = 16
  xlfTan* = 17
  xlfAtan* = 18
  xlfPi* = 19
  xlfSqrt* = 20
  xlfExp* = 21
  xlfLn* = 22
  xlfLog10* = 23
  xlfAbs* = 24
  xlfInt* = 25
  xlfSign* = 26
  xlfRound* = 27
  xlfLookup* = 28
  xlfIndex* = 29
  xlfRept* = 30
  xlfMid* = 31
  xlfLen* = 32
  xlfValue* = 33
  xlfTrue* = 34
  xlfFalse* = 35
  xlfAnd* = 36
  xlfOr* = 37
  xlfNot* = 38
  xlfMod* = 39
  xlfDcount* = 40
  xlfDsum* = 41
  xlfDaverage* = 42
  xlfDmin* = 43
  xlfDmax* = 44
  xlfDstdev* = 45
  xlfVar* = 46
  xlfDvar* = 47
  xlfText* = 48
  xlfLinest* = 49
  xlfTrend* = 50
  xlfLogest* = 51
  xlfGrowth* = 52
  xlfGoto* = 53
  xlfHalt* = 54
  xlfPv* = 56
  xlfFv* = 57
  xlfNper* = 58
  xlfPmt* = 59
  xlfRate* = 60
  xlfMirr* = 61
  xlfIrr* = 62
  xlfRand* = 63
  xlfMatch* = 64
  xlfDate* = 65
  xlfTime* = 66
  xlfDay* = 67
  xlfMonth* = 68
  xlfYear* = 69
  xlfWeekday* = 70
  xlfHour* = 71
  xlfMinute* = 72
  xlfSecond* = 73
  xlfNow* = 74
  xlfAreas* = 75
  xlfRows* = 76
  xlfColumns* = 77
  xlfOffset* = 78
  xlfAbsref* = 79
  xlfRelref* = 80
  xlfArgument* = 81
  xlfSearch* = 82
  xlfTranspose* = 83
  xlfError* = 84
  xlfStep* = 85
  xlfType* = 86
  xlfEcho* = 87
  xlfSetName* = 88
  xlfCaller* = 89
  xlfDeref* = 90
  xlfWindows* = 91
  xlfSeries* = 92
  xlfDocuments* = 93
  xlfActiveCell* = 94
  xlfSelection* = 95
  xlfResult* = 96
  xlfAtan2* = 97
  xlfAsin* = 98
  xlfAcos* = 99
  xlfChoose* = 100
  xlfHlookup* = 101
  xlfVlookup* = 102
  xlfLinks* = 103
  xlfInput* = 104
  xlfIsref* = 105
  xlfGetFormula* = 106
  xlfGetName* = 107
  xlfSetValue* = 108
  xlfLog* = 109
  xlfExec* = 110
  xlfChar* = 111
  xlfLower* = 112
  xlfUpper* = 113
  xlfProper* = 114
  xlfLeft* = 115
  xlfRight* = 116
  xlfExact* = 117
  xlfTrim* = 118
  xlfReplace* = 119
  xlfSubstitute* = 120
  xlfCode* = 121
  xlfNames* = 122
  xlfDirectory* = 123
  xlfFind* = 124
  xlfCell* = 125
  xlfIserr* = 126
  xlfIstext* = 127
  xlfIsnumber* = 128
  xlfIsblank* = 129
  xlfT* = 130
  xlfN* = 131
  xlfFopen* = 132
  xlfFclose* = 133
  xlfFsize* = 134
  xlfFreadln* = 135
  xlfFread* = 136
  xlfFwriteln* = 137
  xlfFwrite* = 138
  xlfFpos* = 139
  xlfDatevalue* = 140
  xlfTimevalue* = 141
  xlfSln* = 142
  xlfSyd* = 143
  xlfDdb* = 144
  xlfGetDef* = 145
  xlfReftext* = 146
  xlfTextref* = 147
  xlfIndirect* = 148
  xlfRegister* = 149
  xlfCall* = 150
  xlfAddBar* = 151
  xlfAddMenu* = 152
  xlfAddCommand* = 153
  xlfEnableCommand* = 154
  xlfCheckCommand* = 155
  xlfRenameCommand* = 156
  xlfShowBar* = 157
  xlfDeleteMenu* = 158
  xlfDeleteCommand* = 159
  xlfGetChartItem* = 160
  xlfDialogBox* = 161
  xlfClean* = 162
  xlfMdeterm* = 163
  xlfMinverse* = 164
  xlfMmult* = 165
  xlfFiles* = 166
  xlfIpmt* = 167
  xlfPpmt* = 168
  xlfCounta* = 169
  xlfCancelKey* = 170
  xlfInitiate* = 175
  xlfRequest* = 176
  xlfPoke* = 177
  xlfExecute* = 178
  xlfTerminate* = 179
  xlfRestart* = 180
  xlfHelp* = 181
  xlfGetBar* = 182
  xlfProduct* = 183
  xlfFact* = 184
  xlfGetCell* = 185
  xlfGetWorkspace* = 186
  xlfGetWindow* = 187
  xlfGetDocument* = 188
  xlfDproduct* = 189
  xlfIsnontext* = 190
  xlfGetNote* = 191
  xlfNote* = 192
  xlfStdevp* = 193
  xlfVarp* = 194
  xlfDstdevp* = 195
  xlfDvarp* = 196
  xlfTrunc* = 197
  xlfIslogical* = 198
  xlfDcounta* = 199
  xlfDeleteBar* = 200
  xlfUnregister* = 201
  xlfUsdollar* = 204
  xlfFindb* = 205
  xlfSearchb* = 206
  xlfReplaceb* = 207
  xlfLeftb* = 208
  xlfRightb* = 209
  xlfMidb* = 210
  xlfLenb* = 211
  xlfRoundup* = 212
  xlfRounddown* = 213
  xlfAsc* = 214
  xlfDbcs* = 215
  xlfRank* = 216
  xlfAddress* = 219
  xlfDays360* = 220
  xlfToday* = 221
  xlfVdb* = 222
  xlfMedian* = 227
  xlfSumproduct* = 228
  xlfSinh* = 229
  xlfCosh* = 230
  xlfTanh* = 231
  xlfAsinh* = 232
  xlfAcosh* = 233
  xlfAtanh* = 234
  xlfDget* = 235
  xlfCreateObject* = 236
  xlfVolatile* = 237
  xlfLastError* = 238
  xlfCustomUndo* = 239
  xlfCustomRepeat* = 240
  xlfFormulaConvert* = 241
  xlfGetLinkInfo* = 242
  xlfTextBox* = 243
  xlfInfo* = 244
  xlfGroup* = 245
  xlfGetObject* = 246
  xlfDb* = 247
  xlfPause* = 248
  xlfResume* = 251
  xlfFrequency* = 252
  xlfAddToolbar* = 253
  xlfDeleteToolbar* = 254
  xlfResetToolbar* = 256
  xlfEvaluate* = 257
  xlfGetToolbar* = 258
  xlfGetTool* = 259
  xlfSpellingCheck* = 260
  xlfErrorType* = 261
  xlfAppTitle* = 262
  xlfWindowTitle* = 263
  xlfSaveToolbar* = 264
  xlfEnableTool* = 265
  xlfPressTool* = 266
  xlfRegisterId* = 267
  xlfGetWorkbook* = 268
  xlfAvedev* = 269
  xlfBetadist* = 270
  xlfGammaln* = 271
  xlfBetainv* = 272
  xlfBinomdist* = 273
  xlfChidist* = 274
  xlfChiinv* = 275
  xlfCombin* = 276
  xlfConfidence* = 277
  xlfCritbinom* = 278
  xlfEven* = 279
  xlfExpondist* = 280
  xlfFdist* = 281
  xlfFinv* = 282
  xlfFisher* = 283
  xlfFisherinv* = 284
  xlfFloor* = 285
  xlfGammadist* = 286
  xlfGammainv* = 287
  xlfCeiling* = 288
  xlfHypgeomdist* = 289
  xlfLognormdist* = 290
  xlfLoginv* = 291
  xlfNegbinomdist* = 292
  xlfNormdist* = 293
  xlfNormsdist* = 294
  xlfNorminv* = 295
  xlfNormsinv* = 296
  xlfStandardize* = 297
  xlfOdd* = 298
  xlfPermut* = 299
  xlfPoisson* = 300
  xlfTdist* = 301
  xlfWeibull* = 302
  xlfSumxmy2* = 303
  xlfSumx2my2* = 304
  xlfSumx2py2* = 305
  xlfChitest* = 306
  xlfCorrel* = 307
  xlfCovar* = 308
  xlfForecast* = 309
  xlfFtest* = 310
  xlfIntercept* = 311
  xlfPearson* = 312
  xlfRsq* = 313
  xlfSteyx* = 314
  xlfSlope* = 315
  xlfTtest* = 316
  xlfProb* = 317
  xlfDevsq* = 318
  xlfGeomean* = 319
  xlfHarmean* = 320
  xlfSumsq* = 321
  xlfKurt* = 322
  xlfSkew* = 323
  xlfZtest* = 324
  xlfLarge* = 325
  xlfSmall* = 326
  xlfQuartile* = 327
  xlfPercentile* = 328
  xlfPercentrank* = 329
  xlfMode* = 330
  xlfTrimmean* = 331
  xlfTinv* = 332
  xlfMovieCommand* = 334
  xlfGetMovie* = 335
  xlfConcatenate* = 336
  xlfPower* = 337
  xlfPivotAddData* = 338
  xlfGetPivotTable* = 339
  xlfGetPivotField* = 340
  xlfGetPivotItem* = 341
  xlfRadians* = 342
  xlfDegrees* = 343
  xlfSubtotal* = 344
  xlfSumif* = 345
  xlfCountif* = 346
  xlfCountblank* = 347
  xlfScenarioGet* = 348
  xlfOptionsListsGet* = 349
  xlfIspmt* = 350
  xlfDatedif* = 351
  xlfDatestring* = 352
  xlfNumberstring* = 353
  xlfRoman* = 354
  xlfOpenDialog* = 355
  xlfSaveDialog* = 356
  xlfViewGet* = 357
  xlfGetpivotdata* = 358
  xlfHyperlink* = 359
  xlfPhonetic* = 360
  xlfAveragea* = 361
  xlfMaxa* = 362
  xlfMina* = 363
  xlfStdevpa* = 364
  xlfVarpa* = 365
  xlfStdeva* = 366
  xlfVara* = 367
  xlfBahttext* = 368
  xlfThaidayofweek* = 369
  xlfThaidigit* = 370
  xlfThaimonthofyear* = 371
  xlfThainumsound* = 372
  xlfThainumstring* = 373
  xlfThaistringlength* = 374
  xlfIsthaidigit* = 375
  xlfRoundbahtdown* = 376
  xlfRoundbahtup* = 377
  xlfThaiyear* = 378
  xlfRtd* = 379
  xlfCubevalue* = 380
  xlfCubemember* = 381
  xlfCubememberproperty* = 382
  xlfCuberankedmember* = 383
  xlfHex2bin* = 384
  xlfHex2dec* = 385
  xlfHex2oct* = 386
  xlfDec2bin* = 387
  xlfDec2hex* = 388
  xlfDec2oct* = 389
  xlfOct2bin* = 390
  xlfOct2hex* = 391
  xlfOct2dec* = 392
  xlfBin2dec* = 393
  xlfBin2oct* = 394
  xlfBin2hex* = 395
  xlfImsub* = 396
  xlfImdiv* = 397
  xlfImpower* = 398
  xlfImabs* = 399
  xlfImsqrt* = 400
  xlfImln* = 401
  xlfImlog2* = 402
  xlfImlog10* = 403
  xlfImsin* = 404
  xlfImcos* = 405
  xlfImexp* = 406
  xlfImargument* = 407
  xlfImconjugate* = 408
  xlfImaginary* = 409
  xlfImreal* = 410
  xlfComplex* = 411
  xlfImsum* = 412
  xlfImproduct* = 413
  xlfSeriessum* = 414
  xlfFactdouble* = 415
  xlfSqrtpi* = 416
  xlfQuotient* = 417
  xlfDelta* = 418
  xlfGestep* = 419
  xlfIseven* = 420
  xlfIsodd* = 421
  xlfMround* = 422
  xlfErf* = 423
  xlfErfc* = 424
  xlfBesselj* = 425
  xlfBesselk* = 426
  xlfBessely* = 427
  xlfBesseli* = 428
  xlfXirr* = 429
  xlfXnpv* = 430
  xlfPricemat* = 431
  xlfYieldmat* = 432
  xlfIntrate* = 433
  xlfReceived* = 434
  xlfDisc* = 435
  xlfPricedisc* = 436
  xlfYielddisc* = 437
  xlfTbilleq* = 438
  xlfTbillprice* = 439
  xlfTbillyield* = 440
  xlfPrice* = 441
  xlfYield* = 442
  xlfDollarde* = 443
  xlfDollarfr* = 444
  xlfNominal* = 445
  xlfEffect* = 446
  xlfCumprinc* = 447
  xlfCumipmt* = 448
  xlfEdate* = 449
  xlfEomonth* = 450
  xlfYearfrac* = 451
  xlfCoupdaybs* = 452
  xlfCoupdays* = 453
  xlfCoupdaysnc* = 454
  xlfCoupncd* = 455
  xlfCoupnum* = 456
  xlfCouppcd* = 457
  xlfDuration* = 458
  xlfMduration* = 459
  xlfOddlprice* = 460
  xlfOddlyield* = 461
  xlfOddfprice* = 462
  xlfOddfyield* = 463
  xlfRandbetween* = 464
  xlfWeeknum* = 465
  xlfAmordegrc* = 466
  xlfAmorlinc* = 467
  xlfConvert* = 468
  xlfAccrint* = 469
  xlfAccrintm* = 470
  xlfWorkday* = 471
  xlfNetworkdays* = 472
  xlfGcd* = 473
  xlfMultinomial* = 474
  xlfLcm* = 475
  xlfFvschedule* = 476
  xlfCubekpimember* = 477
  xlfCubeset* = 478
  xlfCubesetcount* = 479
  xlfIferror* = 480
  xlfCountifs* = 481
  xlfSumifs* = 482
  xlfAverageif* = 483
  xlfAverageifs* = 484
  xlfAggregate* = 485
  xlfBinomdist_2* = 486
  xlfBinominv* = 487
  xlfConfidencenorm* = 488
  xlfConfidencet* = 489
  xlfChisqtest* = 490
  xlfFtest_2* = 491
  xlfCovariancep* = 492
  xlfCovariances* = 493
  xlfExpondist_2* = 494
  xlfGammadist_2* = 495
  xlfGammainv_2* = 496
  xlfModemult* = 497
  xlfModesngl* = 498
  xlfNormdist_2* = 499
  xlfNorminv_2* = 500
  xlfPercentileexc* = 501
  xlfPercentileinc* = 502
  xlfPercentrankexc* = 503
  xlfPercentrankinc* = 504
  xlfPoissondist* = 505
  xlfQuartileexc* = 506
  xlfQuartileinc* = 507
  xlfRankavg* = 508
  xlfRankeq* = 509
  xlfStdevs* = 510
  xlfStdevp_2* = 511
  xlfTdist_2* = 512
  xlfTdist2t* = 513
  xlfTdistrt* = 514
  xlfTinv_2* = 515
  xlfTinv2t* = 516
  xlfVars* = 517
  xlfVarp_2* = 518
  xlfWeibulldist* = 519
  xlfNetworkdaysintl* = 520
  xlfWorkdayintl* = 521
  xlfEcmaceiling* = 522
  xlfIsoceiling* = 523
  xlfBetadist_2* = 525
  xlfBetainv_2* = 526
  xlfChisqdist* = 527
  xlfChisqdistrt* = 528
  xlfChisqinv* = 529
  xlfChisqinvrt* = 530
  xlfFdist_2* = 531
  xlfFdistrt* = 532
  xlfFinv_2* = 533
  xlfFinvrt* = 534
  xlfHypgeomdist_2* = 535
  xlfLognormdist_2* = 536
  xlfLognorminv* = 537
  xlfNegbinomdist_2* = 538
  xlfNormsdist_2* = 539
  xlfNormsinv_2* = 540
  xlfTtest_2* = 541
  xlfZtest_2* = 542
  xlfErfprecise* = 543
  xlfErfcprecise* = 544
  xlfGammalnprecise* = 545
  xlfCeilingprecise* = 546
  xlfFloorprecise* = 547
  xlfAcot* = 548
  xlfAcoth* = 549
  xlfCot* = 550
  xlfCoth* = 551
  xlfCsc* = 552
  xlfCsch* = 553
  xlfSec* = 554
  xlfSech* = 555
  xlfImtan* = 556
  xlfImcot* = 557
  xlfImcsc* = 558
  xlfImcsch* = 559
  xlfImsec* = 560
  xlfImsech* = 561
  xlfBitand* = 562
  xlfBitor* = 563
  xlfBitxor* = 564
  xlfBitlshift* = 565
  xlfBitrshift* = 566
  xlfPermutationa* = 567
  xlfCombina* = 568
  xlfXor* = 569
  xlfPduration* = 570
  xlfBase* = 571
  xlfDecimal* = 572
  xlfDays* = 573
  xlfBinomdistrange* = 574
  xlfGamma* = 575
  xlfSkewp* = 576
  xlfGauss* = 577
  xlfPhi* = 578
  xlfRri* = 579
  xlfUnichar* = 580
  xlfUnicode* = 581
  xlfMunit* = 582
  xlfArabic* = 583
  xlfIsoweeknum* = 584
  xlfNumbervalue* = 585
  xlfSheet* = 586
  xlfSheets* = 587
  xlfFormulatext* = 588
  xlfIsformula* = 589
  xlfIfna* = 590
  xlfCeilingmath* = 591
  xlfFloormath* = 592
  xlfImsinh* = 593
  xlfImcosh* = 594
  xlfFilterxml* = 595
  xlfWebservice* = 596
  xlfEncodeurl* = 597
