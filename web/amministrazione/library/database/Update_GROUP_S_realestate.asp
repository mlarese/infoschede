<!--#INCLUDE FILE="Update__FileHeader.asp" -->
<% '........................................................................................... %>
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<!--#INCLUDE FILE="../ToolsDescrittori.asp" -->
<!--#INCLUDE FILE="../Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../IndexContent/Tools_IndexContent.asp" -->
<%
'...........................................................................................
set index.conn = conn
'...........................................................................................


'*******************************************************************************************
'AGGIORNAMENTO PER ALLINEARE PRESTIGE INTERNATIONAL ( Nicola - 13/02/2014)
'...........................................................................................
if lCase(GetDatabaseName(conn)) = "prestigeinternational" then
	sql = "UPDATE AA_versione SET versione=488"
	CALL DB.Execute(sql, 326)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO PER ALLINEARE VENICE REAL ESTATE ( Nicola - 13/02/2014)
'...........................................................................................
if lCase(GetDatabaseName(conn)) = "venicerealestate" then
	'installa memo2 su venicerealestate.
	sql = Install__MEMO2(conn)
	CALL DB.Execute(sql, 538)
	
	sql = AggiornamentoSpeciale__MEMO2__1(DB, rs, 539)
	CALL DB.Execute(sql, 539)
	
	sql = Aggiornamento__MEMO2__1(conn)
	CALL DB.Execute(sql, 540)
	
	sql = Aggiornamento__MEMO2__2(conn)
	CALL DB.Execute(sql, 541)
	if DB.last_update_executed then
		CALL AggiornamentoSpeciale__MEMO2__2(conn)
	end if
	
	sql = Aggiornamento__MEMO2__3(conn)
	CALL DB.Execute(sql, 542)
	
	sql = Aggiornamento__MEMO2__4(conn)
	CALL DB.Execute(sql, 543)
	if DB.last_update_executed then
		CALL AggiornamentoSpeciale__MEMO2__4(conn)
	end if
	
	sql = Aggiornamento__MEMO2__5(conn)
	CALL DB.Execute(sql, 544)
	if DB.last_update_executed then
		CALL AggiornamentoSpeciale__MEMO2__5(conn)
	end if
	
	sql = Aggiornamento__MEMO2__6(conn)
	CALL DB.Execute(sql, 545)
	
	sql = Aggiornamento__MEMO2__7(conn)
	CALL DB.Execute(sql, 546)
	if DB.last_update_executed then
		CALL AggiornamentoSpeciale__MEMO2__7(conn)
	end if

	sql = Aggiornamento__MEMO2__8(conn)
	CALL DB.Execute(sql, 547)
	if DB.last_update_executed then
		CALL AggiornamentoSpeciale__MEMO2__8(conn)
	end if
	
	sql = Aggiornamento__MEMO2__9(conn)
	CALL DB.Execute(sql, 548)
	
	sql = Aggiornamento__MEMO2__10(conn)
	CALL DB.Execute(sql, 549)
	
	sql = Aggiornamento__MEMO2__11(conn)
	CALL DB.Execute(sql, 550)
	
	sql = Aggiornamento__MEMO2__12(conn)
	CALL DB.Execute(sql, 551)
	
	sql = "UPDATE AA_versione SET versione=486"
	CALL DB.Execute(sql, 552)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO PER ALLINEARE VENICE REAL ESTATE AL MODEL DEL BIPLATFORM ( Luca - 21/02/2017)
'...........................................................................................
if lCase(GetDatabaseName(conn)) = "venicerealestatenew" then
	sql = Install_B2B__version_79()
	CALL DB.Execute(sql, 505)
	
	sql = Indexing_B2B()
	CALL DB.Execute(sql, 506)
	
	sql = Aggiornamento__B2B__1(conn)
	CALL DB.Execute(sql, 507)
	
	sql = Aggiornamento__B2B__2(conn)
	CALL DB.Execute(sql, 508)
	
	sql = Aggiornamento__B2B__3(conn)
	CALL DB.Execute(sql, 509)
	
	sql = Aggiornamento__B2B__4(conn)
	CALL DB.Execute(sql, 510)
	
	sql = Aggiornamento__B2B__5(conn)
	CALL DB.Execute(sql, 511)
	
	sql = Aggiornamento__B2B__6(conn)
	CALL DB.Execute(sql, 512)
	
	sql = Aggiornamento__B2B__7(conn)
	CALL DB.Execute(sql, 513)
	
	sql = Aggiornamento__B2B__8(conn)
	CALL DB.Execute(sql, 514)
	
	sql = Aggiornamento__B2B__9(conn)
	CALL DB.Execute(sql, 515)
	
	sql = Aggiornamento__B2B__10(conn)
	CALL DB.Execute(sql, 516)
	
	sql = Aggiornamento__B2B__11(conn)
	CALL DB.Execute(sql, 517)
	
	sql = Aggiornamento__B2B__12(conn)
	CALL DB.Execute(sql, 518)
	
	sql = Aggiornamento__B2B__13(conn)
	CALL DB.Execute(sql, 519)
	
	sql = Aggiornamento__B2B__14(conn)
	CALL DB.Execute(sql, 520)
	
	sql = Aggiornamento__B2B__15(conn)
	CALL DB.Execute(sql, 521)
	
	sql = Aggiornamento__B2B__16(conn)
	CALL DB.Execute(sql, 522)
	
	sql = Aggiornamento__B2B__17(conn)
	CALL DB.Execute(sql, 523)
	
	sql = Aggiornamento__B2B__18(conn)
	CALL DB.Execute(sql, 524)
	
	sql = Aggiornamento__B2B__19(conn)
	CALL DB.Execute(sql, 525)
	
	sql = Aggiornamento__B2B__20(conn)
	CALL DB.Execute(sql, 526)
	
	sql = Aggiornamento__B2B__21(conn)
	CALL DB.Execute(sql, 527)
	
	sql = Aggiornamento__B2B__22(conn)
	CALL DB.Execute(sql, 528)
	
	sql = Aggiornamento__B2B__23(conn)
	CALL DB.Execute(sql, 529)
	
	sql = Aggiornamento__B2B__24(conn)
	CALL DB.Execute(sql, 530)
	
	sql = Aggiornamento__B2B__25(conn)
	CALL DB.Execute(sql, 531)
	
	sql = Aggiornamento__B2B__26(conn)
	CALL DB.Execute(sql, 532)
	
	sql = Aggiornamento__B2B__27(conn)
	CALL DB.Execute(sql, 533)
	
	sql = Aggiornamento__B2B__28(conn)
	CALL DB.Execute(sql, 534)
	
	sql = Aggiornamento__B2B__29(conn)
	CALL DB.Execute(sql, 535)
	
	sql = Aggiornamento__B2B__30(conn)
	CALL DB.Execute(sql, 536)
	
	sql = Aggiornamento__B2B__31(conn)
	CALL DB.Execute(sql, 537)
	
	sql = Aggiornamento__B2B__32(conn)
	CALL DB.Execute(sql, 538)
	
	sql = Aggiornamento__B2B__33(conn)
	CALL DB.Execute(sql, 539)
	
	sql = Aggiornamento__B2B__34(conn)
	CALL DB.Execute(sql, 540)
	
	sql = Aggiornamento__B2B__35(conn)
	CALL DB.Execute(sql, 541)
	
	sql = Aggiornamento__B2B__36(conn)
	CALL DB.Execute(sql, 542)
	
	sql = Aggiornamento__B2B__37(conn)
	CALL DB.Execute(sql, 543)
	
	sql = Aggiornamento__B2B__38(conn)
	CALL DB.Execute(sql, 544)
	
	sql = Aggiornamento__B2B__39(conn)
	CALL DB.Execute(sql, 545)
	
	sql = Aggiornamento__B2B__40(conn)
	CALL DB.Execute(sql, 546)
	
	sql = Aggiornamento__B2B__41(conn)
	CALL DB.Execute(sql, 547)
	
	sql = Aggiornamento__B2B__42(conn)
	CALL DB.Execute(sql, 548)
	
	sql = Aggiornamento__B2B__43(conn)
	CALL DB.Execute(sql, 549)
	
	sql = Aggiornamento__B2B__44(conn)
	CALL DB.Execute(sql, 550)
	
	sql = Aggiornamento__B2B__45(conn)
	CALL DB.Execute(sql, 551)
	
	sql = Aggiornamento__B2B__46(conn)
	CALL DB.Execute(sql, 552)
	
	sql = Aggiornamento__B2B__47(conn)
	CALL DB.Execute(sql, 553)
	
	sql = Aggiornamento__B2B__48(conn)
	CALL DB.Execute(sql, 554)
	
	sql = Aggiornamento__B2B__49(conn)
	CALL DB.Execute(sql, 555)
	
	sql = Aggiornamento__B2B__50(conn)
	CALL DB.Execute(sql, 556)
	
	sql = Aggiornamento__B2B__51(conn)
	CALL DB.Execute(sql, 557)
	
	sql = Aggiornamento__B2B__52(conn)
	CALL DB.Execute(sql, 558)
	
	sql = Aggiornamento__B2B__53(conn)
	CALL DB.Execute(sql, 559)
	
	sql = Aggiornamento__B2B__54(conn)
	CALL DB.Execute(sql, 560)
	
	sql = Aggiornamento__B2B__55(conn)
	CALL DB.Execute(sql, 561)
	
	sql = Aggiornamento__B2B__56(conn)
	CALL DB.Execute(sql, 562)
	
	sql = Aggiornamento__B2B__57(conn)
	CALL DB.Execute(sql, 563)
	
	sql = Aggiornamento__B2B__58(conn)
	CALL DB.Execute(sql, 564)
	
	sql = Aggiornamento__B2B__59(conn)
	CALL DB.Execute(sql, 565)
	
	sql = Aggiornamento__B2B__60(conn)
	CALL DB.Execute(sql, 566)
	
	sql = Aggiornamento__B2B__61(conn)
	CALL DB.Execute(sql, 567)
	
	sql = Aggiornamento__B2B__62(conn)
	CALL DB.Execute(sql, 568)
	
	sql = Aggiornamento__B2B__63(conn)
	CALL DB.Execute(sql, 569)
	
	sql = Aggiornamento__B2B__64(conn)
	CALL DB.Execute(sql, 570)
	
	sql = Aggiornamento__B2B__65(conn)
	CALL DB.Execute(sql, 571)
	
	sql = Aggiornamento__B2B__66(conn)
	CALL DB.Execute(sql, 572)
	
	sql = Aggiornamento__B2B__67(conn)
	CALL DB.Execute(sql, 573)
	
	sql = Aggiornamento__B2B__68(conn)
	CALL DB.Execute(sql, 574)
	
	sql = Aggiornamento__B2B__69(conn)
	CALL DB.Execute(sql, 575)
	
	sql = Aggiornamento__B2B__70(conn)
	CALL DB.Execute(sql, 576)
	
	sql = Aggiornamento__B2B__71(conn)
	CALL DB.Execute(sql, 577)
	
	sql = Aggiornamento__B2B__72(conn)
	CALL DB.Execute(sql, 578)
	
	sql = Aggiornamento__B2B__73(conn)
	CALL DB.Execute(sql, 579)
	
	sql = Aggiornamento__B2B__74(conn)
	CALL DB.Execute(sql, 580)
	
	sql = Aggiornamento__B2B__75(conn)
	CALL DB.Execute(sql, 581)
	
	sql = Aggiornamento__B2B__76(conn)
	CALL DB.Execute(sql, 582)
	
	sql = Aggiornamento__B2B__77(conn)
	CALL DB.Execute(sql, 583)
	
	sql = Aggiornamento__B2B__78(conn)
	CALL DB.Execute(sql, 584)
	
	sql = Aggiornamento__B2B__79(conn)
	CALL DB.Execute(sql, 585)
	
	sql = Aggiornamento__B2B__80(conn)
	CALL DB.Execute(sql, 586)
	
	sql = Aggiornamento__B2B__81(conn)
	CALL DB.Execute(sql, 587)
	
	sql = Aggiornamento__B2B__82(conn)
	CALL DB.Execute(sql, 588)
	
	sql = Aggiornamento__B2B__83(conn)
	CALL DB.Execute(sql, 589)
	
	sql = Aggiornamento__B2B__84(conn)
	CALL DB.Execute(sql, 590)
	
	sql = Aggiornamento__B2B__85(conn)
	CALL DB.Execute(sql, 591)
	
	sql = Aggiornamento__B2B__86(conn)
	CALL DB.Execute(sql, 592)
	
	sql = Aggiornamento__B2B__87(conn)
	CALL DB.Execute(sql, 593)
	
	sql = Aggiornamento__B2B__88(conn)
	CALL DB.Execute(sql, 594)
	
	sql = Aggiornamento__B2B__89(conn)
	CALL DB.Execute(sql, 595)
	
	sql = Aggiornamento__B2B__90(conn)
	CALL DB.Execute(sql, 596)
	
	sql = Aggiornamento__B2B__91(conn)
	CALL DB.Execute(sql, 597)
	
	sql = Aggiornamento__B2B__92(conn)
	CALL DB.Execute(sql, 598)
	
	sql = Aggiornamento__B2B__93(conn)
	CALL DB.Execute(sql, 599)
	
	sql = Aggiornamento__B2B__94(conn)
	CALL DB.Execute(sql, 600)
	
	sql = Aggiornamento__B2B__95(conn)
	CALL DB.Execute(sql, 601)
	
	sql = Aggiornamento__B2B__96(conn)
	CALL DB.Execute(sql, 602)
	
	sql = Aggiornamento__B2B__97(conn)
	CALL DB.Execute(sql, 603)
	
	sql = Aggiornamento__B2B__98(conn)
	CALL DB.Execute(sql, 604)
	
	sql = Aggiornamento__B2B__99(conn)
	CALL DB.Execute(sql, 605)
	
	sql = Aggiornamento__B2B__100(conn)
	CALL DB.Execute(sql, 606)
	
	sql = Aggiornamento__B2B__101(conn)
	CALL DB.Execute(sql, 607)
	
	sql = Aggiornamento__B2B__102(conn)
	CALL DB.Execute(sql, 608)
	
	sql = Aggiornamento__B2B__103(conn)
	CALL DB.Execute(sql, 609)
	
	sql = Aggiornamento__B2B__104(conn)
	CALL DB.Execute(sql, 610)
	
	sql = Aggiornamento__B2B__105(conn)
	CALL DB.Execute(sql, 611)
	
	sql = Aggiornamento__B2B__106(conn)
	CALL DB.Execute(sql, 612)
	
	sql = Aggiornamento__B2B__107(conn)
	CALL DB.Execute(sql, 613)
	
	sql = Aggiornamento__B2B__108(conn)
	CALL DB.Execute(sql, 614)
	
	sql = Aggiornamento__B2B__109(conn)
	CALL DB.Execute(sql, 615)
	
	sql = Aggiornamento__B2B__110(conn)
	CALL DB.Execute(sql, 616)
	
	sql = Aggiornamento__B2B__111(conn)
	CALL DB.Execute(sql, 617)
	
	sql = Aggiornamento__B2B__112(conn)
	CALL DB.Execute(sql, 618)
	
	sql = Aggiornamento__B2B__113(conn)
	CALL DB.Execute(sql, 619)
	
	sql = Aggiornamento__B2B__114(conn)
	CALL DB.Execute(sql, 620)
	
	sql = Aggiornamento__B2B__115(conn)
	CALL DB.Execute(sql, 621)
	
	sql = Aggiornamento__B2B__116(conn)
	CALL DB.Execute(sql, 622)
	
	sql = Aggiornamento__B2B__117(conn)
	CALL DB.Execute(sql, 623)
	
	sql = Aggiornamento__B2B__118(conn)
	CALL DB.Execute(sql, 624)
	
	sql = Aggiornamento__B2B__119(conn)
	CALL DB.Execute(sql, 625)
	
	sql = Aggiornamento__B2B__120(conn)
	CALL DB.Execute(sql, 626)
	
	sql = Aggiornamento__B2B__121(conn)
	CALL DB.Execute(sql, 627)
	
	sql = Aggiornamento__B2B__122(conn)
	CALL DB.Execute(sql, 628)
	
	sql = Aggiornamento__B2B__123(conn)
	CALL DB.Execute(sql, 629)
	
	sql = Aggiornamento__B2B__124(conn)
	CALL DB.Execute(sql, 630)
	
	sql = Aggiornamento__B2B__125(conn)
	CALL DB.Execute(sql, 631)
	
	sql = Aggiornamento__B2B__126(conn)
	CALL DB.Execute(sql, 632)
	
	sql = Aggiornamento__B2B__127(conn)
	CALL DB.Execute(sql, 633)
	
	sql = Aggiornamento__B2B__128(conn)
	CALL DB.Execute(sql, 634)
	
	sql = Aggiornamento__B2B__129(conn)
	CALL DB.Execute(sql, 635)
	
	sql = Aggiornamento__B2B__130(conn)
	CALL DB.Execute(sql, 636)
	
	sql = Aggiornamento__B2B__131(conn)
	CALL DB.Execute(sql, 637)
	
	sql = Aggiornamento__B2B__132(conn)
	CALL DB.Execute(sql, 638)
	
	sql = Aggiornamento__B2B__133(conn)
	CALL DB.Execute(sql, 639)
	
	sql = Aggiornamento__B2B__134(conn)
	CALL DB.Execute(sql, 640)
	
	sql = Aggiornamento__B2B__135(conn)
	CALL DB.Execute(sql, 641)
	
	sql = Aggiornamento__B2B__136(conn)
	CALL DB.Execute(sql, 642)
	
	sql = Aggiornamento__B2B__137(conn)
	CALL DB.Execute(sql, 643)
	
	sql = Aggiornamento__B2B__138(conn)
	CALL DB.Execute(sql, 644)
	
	sql = Aggiornamento__B2B__139(conn)
	CALL DB.Execute(sql, 645)
	
	sql = Aggiornamento__B2B__140(conn)
	CALL DB.Execute(sql, 646)
	
	sql = Aggiornamento__B2B__141(conn)
	CALL DB.Execute(sql, 647)
	
	sql = Aggiornamento__B2B__142(conn)
	CALL DB.Execute(sql, 648)
	
	sql = Aggiornamento__B2B__143(conn)
	CALL DB.Execute(sql, 649)
	
	sql = Aggiornamento__B2B__144(conn)
	CALL DB.Execute(sql, 650)
	
	sql = Aggiornamento__B2B__145(conn)
	CALL DB.Execute(sql, 651)
	
	sql = Aggiornamento__B2B__146(conn)
	CALL DB.Execute(sql, 652)
	
	sql = Aggiornamento__B2B__147(conn)
	CALL DB.Execute(sql, 653)
	
	sql = Aggiornamento__B2B__148(conn)
	CALL DB.Execute(sql, 654)
	
	sql = Aggiornamento__B2B__149(conn)
	CALL DB.Execute(sql, 655)
	
	sql = Aggiornamento__B2B__150(conn)
	CALL DB.Execute(sql, 656)
	
	sql = Aggiornamento__B2B__151(conn)
	CALL DB.Execute(sql, 657)
	
	sql = Aggiornamento__B2B__152(conn)
	CALL DB.Execute(sql, 658)
	
	sql = Aggiornamento__B2B__153(conn)
	CALL DB.Execute(sql, 659)
	
	sql = Aggiornamento__B2B__154(conn)
	CALL DB.Execute(sql, 660)
	
	sql = Aggiornamento__B2B__155(conn)
	CALL DB.Execute(sql, 661)
	
	sql = Aggiornamento__B2B__156(conn)
	CALL DB.Execute(sql, 662)
	
	sql = Aggiornamento__B2B__157(conn)
	CALL DB.Execute(sql, 663)
	
	sql = Aggiornamento__B2B__158(conn)
	CALL DB.Execute(sql, 664)
	
	sql = Aggiornamento__B2B__159(conn)
	CALL DB.Execute(sql, 665)
	
	sql = Aggiornamento__B2B__160(conn)
	CALL DB.Execute(sql, 666)
	
	sql = Aggiornamento__B2B__161(conn)
	CALL DB.Execute(sql, 667)
	
	sql = Aggiornamento__B2B__162(conn)
	CALL DB.Execute(sql, 668)
	
	sql = Aggiornamento__B2B__163(conn)
	CALL DB.Execute(sql, 669)
	
	sql = Aggiornamento__B2B__164(conn)
	CALL DB.Execute(sql, 670)
	
	sql = Aggiornamento__B2B__165(conn)
	CALL DB.Execute(sql, 671)
	
	sql = Aggiornamento__B2B__166(conn)
	CALL DB.Execute(sql, 672)
	
	sql = Aggiornamento__B2B__167(conn)
	CALL DB.Execute(sql, 673)
	
	sql = Aggiornamento__B2B__168(conn)
	CALL DB.Execute(sql, 674)
	
	sql = Aggiornamento__B2B__169(conn)
	CALL DB.Execute(sql, 675)
	
	sql = Aggiornamento__B2B__170(conn)
	CALL DB.Execute(sql, 676)
	
	sql = Aggiornamento__B2B__171(conn)
	CALL DB.Execute(sql, 677)
	
	sql = Aggiornamento__B2B__172(conn)
	CALL DB.Execute(sql, 678)
	
	sql = Aggiornamento__B2B__173(conn)
	CALL DB.Execute(sql, 679)
	
	sql = Aggiornamento__B2B__174(conn)
	CALL DB.Execute(sql, 680)
	
	sql = Aggiornamento__B2B__175(conn)
	CALL DB.Execute(sql, 681)
	
	sql = Aggiornamento__B2B__176(conn)
	CALL DB.Execute(sql, 682)
	
	sql = Aggiornamento__B2B__177(conn)
	CALL DB.Execute(sql, 683)
	
	sql = Aggiornamento__B2B__178(conn)
	CALL DB.Execute(sql, 684)
	
	sql = Aggiornamento__B2B__179(conn)
	CALL DB.Execute(sql, 685)
	
	sql = Aggiornamento__B2B__180(conn)
	CALL DB.Execute(sql, 686)
	
	sql = Aggiornamento__B2B__181(conn)
	CALL DB.Execute(sql, 687)
	
	sql = Aggiornamento__B2B__182(conn)
	CALL DB.Execute(sql, 688)
	
	sql = Aggiornamento__B2B__183(conn)
	CALL DB.Execute(sql, 689)
	
	sql = Aggiornamento__B2B__184(conn)
	CALL DB.Execute(sql, 690)
	
	sql = Aggiornamento__B2B__185(conn)
	CALL DB.Execute(sql, 691)
	
	sql = Aggiornamento__B2B__186(conn)
	CALL DB.Execute(sql, 692)
	
	sql = Aggiornamento__B2B__187(conn)
	CALL DB.Execute(sql, 693)
	
	sql = Aggiornamento__B2B__188(conn)
	CALL DB.Execute(sql, 694)
	
	sql = Aggiornamento__B2B__189(conn)
	CALL DB.Execute(sql, 695)
	
	sql = Aggiornamento__B2B__190(conn)
	CALL DB.Execute(sql, 696)
	
	sql = Aggiornamento__B2B__191(conn)
	CALL DB.Execute(sql, 697)
	
	sql = Aggiornamento__B2B__192(conn)
	CALL DB.Execute(sql, 698)
	
	sql = Aggiornamento__B2B__193(conn)
	CALL DB.Execute(sql, 699)
	
	sql = Aggiornamento__B2B__194(conn)
	CALL DB.Execute(sql, 700)
	
	sql = Aggiornamento__B2B__195(conn)
	CALL DB.Execute(sql, 701)
	
	sql = Aggiornamento__B2B__196(conn)
	CALL DB.Execute(sql, 702)
	
	sql = Aggiornamento__B2B__197(conn)
	CALL DB.Execute(sql, 703)
	
	sql = Aggiornamento__B2B__198(conn)
	CALL DB.Execute(sql, 704)
	
	sql = Aggiornamento__B2B__199(conn)
	CALL DB.Execute(sql, 705)
	
	sql = Aggiornamento__B2B__200(conn)
	CALL DB.Execute(sql, 706)
	
	sql = Aggiornamento__B2B__201(conn)
	CALL DB.Execute(sql, 707)
	
	sql = Aggiornamento__B2B__202(conn)
	CALL DB.Execute(sql, 708)
	
	sql = Aggiornamento__B2B__203(conn)
	CALL DB.Execute(sql, 709)
	
	sql = Aggiornamento__B2B__204(conn)
	CALL DB.Execute(sql, 710)
	
	sql = Aggiornamento__B2B__205(conn)
	CALL DB.Execute(sql, 711)
	
	sql = Aggiornamento__B2B__206(conn)
	CALL DB.Execute(sql, 712)
	
	sql = Aggiornamento__B2B__207(conn)
	CALL DB.Execute(sql, 713)
	
	sql = Aggiornamento__B2B__208(conn)
	CALL DB.Execute(sql, 714)
	
	sql = Aggiornamento__B2B__209(conn)
	CALL DB.Execute(sql, 715)
	
	sql = Aggiornamento__B2B__210(conn)
	CALL DB.Execute(sql, 716)
	
	sql = Aggiornamento__B2B__211(conn)
	CALL DB.Execute(sql, 717)
	
	sql = Aggiornamento__B2B__212(conn)
	CALL DB.Execute(sql, 718)
	
	sql = Aggiornamento__B2B__213(conn)
	CALL DB.Execute(sql, 719)
	
	sql = Aggiornamento__B2B__214(conn)
	CALL DB.Execute(sql, 720)
	
	sql = Aggiornamento__B2B__215(conn)
	CALL DB.Execute(sql, 721)
	
	sql = Aggiornamento__B2B__216(conn)
	CALL DB.Execute(sql, 722)
	
	sql = Aggiornamento__B2B__217(conn)
	CALL DB.Execute(sql, 723)
	
	sql = Aggiornamento__B2B__218(conn)
	CALL DB.Execute(sql, 724)
	
	sql = Aggiornamento__B2B__219(conn)
	CALL DB.Execute(sql, 725)
	
	sql = Aggiornamento__B2B__220(conn)
	CALL DB.Execute(sql, 726)
	
	sql = Aggiornamento__B2B__221(conn)
	CALL DB.Execute(sql, 727)
	
	sql = Aggiornamento__B2B__222(conn)
	CALL DB.Execute(sql, 728)
	
	sql = Aggiornamento__B2B__223(conn)
	CALL DB.Execute(sql, 729)
	
	sql = Aggiornamento__B2B__224(conn)
	CALL DB.Execute(sql, 730)
	
	sql = Aggiornamento__B2B__225(conn)
	CALL DB.Execute(sql, 731)
	
	sql = Aggiornamento__B2B__226(conn)
	CALL DB.Execute(sql, 732)
	
	sql = Aggiornamento__B2B__227(conn)
	CALL DB.Execute(sql, 733)
	
	sql = "UPDATE AA_versione SET versione=504"
	CALL DB.Execute(sql, 734)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO PER ALLINEARE CASAVENEZIA ( Nicola - 13/02/2014)
'...........................................................................................
if lCase(GetDatabaseName(conn)) = "casavenezia" then
	'allinea database casavenezia con aggiornamenti
	sql = Aggiornamento__FRAMEWORK_CORE__220(conn)
	CALL DB.ProtectedExecuteRebuild(sql, 276, false, true)

	sql = Aggiornamento__FRAMEWORK_CORE__221(conn)
	CALL DB.ProtectedExecuteRebuild(sql, 277, false, true)
	if DB.last_update_executed then
		CALL AggiornamentoSpeciale__FRAMEWORK_CORE__221(conn)
	end if

	sql = Aggiornamento__FRAMEWORK_CORE__222(conn)
	CALL DB.Execute(sql, 278)
	if DB.last_update_executed then
		CALL AggiornamentoSpeciale__FRAMEWORK_CORE__222(conn)
	end if

	sql = Aggiornamento__FRAMEWORK_CORE__223(conn)
	CALL DB.ProtectedExecuteRebuild(sql, 279, false, true)

	'installa memo2 su casavenezia.
	sql = Install__MEMO2(conn)
	CALL DB.Execute(sql, 280)
	
	sql = AggiornamentoSpeciale__MEMO2__1(DB, rs, 281)
	CALL DB.Execute(sql, 281)
	
	sql = Aggiornamento__MEMO2__1(conn)
	CALL DB.Execute(sql, 282)
	
	sql = Aggiornamento__MEMO2__2(conn)
	CALL DB.Execute(sql, 283)
	if DB.last_update_executed then
		CALL AggiornamentoSpeciale__MEMO2__2(conn)
	end if
	
	sql = Aggiornamento__MEMO2__3(conn)
	CALL DB.Execute(sql, 284)
	
	sql = Aggiornamento__MEMO2__4(conn)
	CALL DB.Execute(sql, 285)
	if DB.last_update_executed then
		CALL AggiornamentoSpeciale__MEMO2__4(conn)
	end if
	
	sql = Aggiornamento__MEMO2__5(conn)
	CALL DB.Execute(sql, 286)
	if DB.last_update_executed then
		CALL AggiornamentoSpeciale__MEMO2__5(conn)
	end if
	
	sql = Aggiornamento__MEMO2__6(conn)
	CALL DB.Execute(sql, 287)
	
	sql = Aggiornamento__MEMO2__7(conn)
	CALL DB.Execute(sql, 288)
	if DB.last_update_executed then
		CALL AggiornamentoSpeciale__MEMO2__7(conn)
	end if

	sql = Aggiornamento__MEMO2__8(conn)
	CALL DB.Execute(sql, 289)
	if DB.last_update_executed then
		CALL AggiornamentoSpeciale__MEMO2__8(conn)
	end if
	
	sql = Aggiornamento__MEMO2__9(conn)
	CALL DB.Execute(sql, 290)
	
	sql = Aggiornamento__MEMO2__10(conn)
	CALL DB.Execute(sql, 291)
	
	sql = Aggiornamento__MEMO2__11(conn)
	CALL DB.Execute(sql, 292)
	
	sql = "UPDATE AA_versione SET versione=471"
	CALL DB.Execute(sql, 293)
end if
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO 377
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__203(conn)
CALL DB.Execute(sql, 377)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__203(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 378
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__204(conn)
CALL DB.Execute(sql, 378)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__204(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 379
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__205(conn)
CALL DB.Execute(sql, 379)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 380
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__206(conn)
CALL DB.Execute(sql, 380)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(380)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 381
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__207(conn)
CALL DB.Execute(sql, 381)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(381)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 382
'...........................................................................................
sql = Aggiornamento__GUESTBOOK__2(conn)
CALL DB.Execute(sql, 382)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 383
'...........................................................................................
sql = Aggiornamento__GUESTBOOK__3(conn)
CALL DB.Execute(sql, 383)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__GUESTBOOK__3(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 384
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__208(conn)
CALL DB.Execute(sql, 384)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 385
'...........................................................................................
sql = Aggiornamento__MEMO2__11(conn)
CALL DB.Execute(sql, 385)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 386
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__209(conn)
CALL DB.Execute(sql, 386)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__209(conn)
end if
'*******************************************************************************************
'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(386)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 387
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__210(conn)
CALL DB.Execute(sql, 387)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 388
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__211(conn)
CALL DB.Execute(sql, 388)
'*******************************************************************************************
'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(388)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 389
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__212(conn)
CALL DB.Execute(sql, 389)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 390
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__213(conn)
CALL DB.Execute(sql, 390)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 391
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__214(conn)
CALL DB.ProtectedExecuteRebuild(sql, 391, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 392
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__215(conn)
CALL DB.ProtectedExecuteRebuild(sql, 392, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 393
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__216(conn)
CALL DB.ProtectedExecuteRebuild(sql, 393, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 394
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__217(conn)
CALL DB.Execute(sql, 394)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__217(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 395
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__218(conn)
CALL DB.ProtectedExecuteRebuild(sql, 395, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 396
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__219(conn)
CALL DB.ProtectedExecuteRebuild(sql, 396, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 397
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__220(conn)
CALL DB.ProtectedExecuteRebuild(sql, 397, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 398
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__221(conn)
CALL DB.ProtectedExecuteRebuild(sql, 398, false, true)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__221(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 399
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__222(conn)
CALL DB.Execute(sql, 399)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__222(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 400
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__223(conn)
CALL DB.ProtectedExecuteRebuild(sql, 400, false, true)
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO 401
'...........................................................................................
sql = Install__REALESTATE(conn)
CALL DB.ProtectedExecuteRebuild(sql, 401, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 402
'...........................................................................................
sql = Aggiornamento__REALESTATE__1(conn)
CALL DB.ProtectedExecuteRebuild(sql, 402, false, true)
'*******************************************************************************************
'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 403
'...........................................................................................
sql = Aggiornamento__REALESTATE__2(conn)
CALL DB.ProtectedExecuteRebuild(sql, 403, false, true)
'*******************************************************************************************
'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 404
'...........................................................................................
sql = Aggiornamento__REALESTATE__3(conn)
CALL DB.ProtectedExecuteRebuild(sql, 404, false, true)
'*******************************************************************************************
'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 405
'...........................................................................................
sql = Aggiornamento__REALESTATE__4(conn)
CALL DB.ProtectedExecuteRebuild(sql, 405, false, true)
'*******************************************************************************************
'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 406
'...........................................................................................
sql = Aggiornamento__REALESTATE__5(conn)
CALL DB.ProtectedExecuteRebuild(sql, 406, false, true)
'*******************************************************************************************
'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 407
'...........................................................................................
sql = Aggiornamento__REALESTATE__6(conn)
CALL DB.ProtectedExecuteRebuild(sql, 407, false, true)
'*******************************************************************************************
'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 408
'...........................................................................................
sql = Aggiornamento__REALESTATE__7(conn)
CALL DB.ProtectedExecuteRebuild(sql, 408, false, true)
'*******************************************************************************************
'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 409
'...........................................................................................
sql = Aggiornamento__REALESTATE__8(conn)
CALL DB.ProtectedExecuteRebuild(sql, 409, false, true)
'*******************************************************************************************
'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 410
'...........................................................................................
sql = Aggiornamento__REALESTATE__9(conn)
CALL DB.ProtectedExecuteRebuild(sql, 410, false, true)
'*******************************************************************************************
'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 411
'...........................................................................................
sql = Aggiornamento__REALESTATE__10(conn)
CALL DB.ProtectedExecuteRebuild(sql, 411, false, true)
'*******************************************************************************************
'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 412
'...........................................................................................
sql = Aggiornamento__REALESTATE__12(conn)
CALL DB.ProtectedExecuteRebuild(sql, 412, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 413
'...........................................................................................
sql = Aggiornamento__REALESTATE__13(conn)
CALL DB.ProtectedExecuteRebuild(sql, 413, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 414
'...........................................................................................
sql = Aggiornamento__REALESTATE__14(conn)
CALL DB.ProtectedExecuteRebuild(sql, 414, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 415
'...........................................................................................
sql = Aggiornamento__REALESTATE__15(conn)
CALL DB.ProtectedExecuteRebuild(sql, 415, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 416
'...........................................................................................
sql = Aggiornamento__REALESTATE__16(conn)
CALL DB.ProtectedExecuteRebuild(sql, 416, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 417
'...........................................................................................
sql = Aggiornamento__REALESTATE__17(conn)
CALL DB.ProtectedExecuteRebuild(sql, 417, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 418
'...........................................................................................
sql = Aggiornamento__REALESTATE__18(conn)
CALL DB.ProtectedExecuteRebuild(sql, 418, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 419
'...........................................................................................
sql = Aggiornamento__REALESTATE__19(conn)
CALL DB.ProtectedExecuteRebuild(sql, 419, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 420
'...........................................................................................
sql = Aggiornamento__REALESTATE__20(conn)
CALL DB.ProtectedExecuteRebuild(sql, 420, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 421
'...........................................................................................
sql = Aggiornamento__REALESTATE__21(conn)
CALL DB.ProtectedExecuteRebuild(sql, 421, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 422
'...........................................................................................
sql = Aggiornamento__REALESTATE__22(conn)
CALL DB.ProtectedExecuteRebuild(sql, 422, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 423
'...........................................................................................
sql = Aggiornamento__REALESTATE__23(conn)
CALL DB.ProtectedExecuteRebuild(sql, 423, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 424
'...........................................................................................
sql = Aggiornamento__REALESTATE__24(conn)
CALL DB.ProtectedExecuteRebuild(sql, 424, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 425
'...........................................................................................
sql = Aggiornamento__REALESTATE__25(conn)
CALL DB.ProtectedExecuteRebuild(sql, 425, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 426
'...........................................................................................
sql = Aggiornamento__REALESTATE__26(conn)
CALL DB.ProtectedExecuteRebuild(sql, 426, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 427
'...........................................................................................
sql = Aggiornamento__REALESTATE__27(conn)
CALL DB.ProtectedExecuteRebuild(sql, 427, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 428
'...........................................................................................
sql = Aggiornamento__REALESTATE__28(conn)
CALL DB.ProtectedExecuteRebuild(sql, 428, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 429
'...........................................................................................
sql = Aggiornamento__REALESTATE__29(conn)
CALL DB.ProtectedExecuteRebuild(sql, 429, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 430
'...........................................................................................
sql = Aggiornamento__REALESTATE__30(conn)
CALL DB.ProtectedExecuteRebuild(sql, 430, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 431
'...........................................................................................
sql = Aggiornamento__REALESTATE__31(conn)
CALL DB.ProtectedExecuteRebuild(sql, 431, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 432
'...........................................................................................
sql = Aggiornamento__REALESTATE__32(conn)
CALL DB.ProtectedExecuteRebuild(sql, 432, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 433
'...........................................................................................
sql = Aggiornamento__REALESTATE__33(conn)
CALL DB.ProtectedExecuteRebuild(sql, 433, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 434
'...........................................................................................
sql = Aggiornamento__REALESTATE__34(conn)
CALL DB.ProtectedExecuteRebuild(sql, 434, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 435
'...........................................................................................
sql = Aggiornamento__REALESTATE__35(conn)
CALL DB.ProtectedExecuteRebuild(sql, 435, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 436
'...........................................................................................
sql = Aggiornamento__REALESTATE__36(conn, "ru") & ";" & _
	  Aggiornamento__REALESTATE__36(conn, "cn") & ";" & _
	  Aggiornamento__REALESTATE__36(conn, "pt")
CALL DB.ProtectedExecuteRebuild(sql, 436, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 437
'...........................................................................................
sql = Aggiornamento__REALESTATE__11(conn)
CALL DB.ProtectedExecuteRebuild(sql, 437, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 438
'...........................................................................................
sql = Aggiornamento__REALESTATE__37(conn)
CALL DB.ProtectedExecuteRebuild(sql, 438, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 439
'...........................................................................................
sql = Aggiornamento__REALESTATE__38(conn)
CALL DB.ProtectedExecuteRebuild(sql, 439, false, true)
'*******************************************************************************************
'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(439)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 440
'...........................................................................................
sql = Aggiornamento__REALESTATE__39(conn)
CALL DB.Execute(sql, 440)
'*******************************************************************************************
'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(440)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 441
'...........................................................................................
sql = Aggiornamento__REALESTATE__40(conn)
CALL DB.Execute(sql, 441)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__40(conn)
end if
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 442
'...........................................................................................
sql = Aggiornamento__REALESTATE__41(conn)
CALL DB.Execute(sql, 442)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 443
'...........................................................................................
sql = Aggiornamento__REALESTATE__42(conn)
CALL DB.Execute(sql, 443)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__42(conn)
end if
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 444
'...........................................................................................
sql = Aggiornamento__REALESTATE__43(conn)
CALL DB.Execute(sql, 444)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 445
'...........................................................................................
sql = Aggiornamento__REALESTATE__44(conn)
CALL DB.Execute(sql, 445)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__44(conn)
end if
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 446
'...........................................................................................
sql = Aggiornamento__REALESTATE__45(conn)
CALL DB.Execute(sql, 446)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 447
'...........................................................................................
sql = Aggiornamento__REALESTATE__46(conn)
CALL DB.Execute(sql, 447)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__46(conn)
end if
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 448
'...........................................................................................
sql = Aggiornamento__REALESTATE__47(conn)
CALL DB.Execute(sql, 448)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 449
'...........................................................................................
sql = Aggiornamento__REALESTATE__48(conn)
CALL DB.Execute(sql, 449)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 450
'...........................................................................................
sql = Aggiornamento__REALESTATE__49(conn)
CALL DB.Execute(sql, 450)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 451
'...........................................................................................
sql = Aggiornamento__REALESTATE__50(conn)
CALL DB.Execute(sql, 451)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 452
'...........................................................................................
sql = Aggiornamento__REALESTATE__51(conn)
CALL DB.Execute(sql, 452)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__51(conn)
end if
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 433
'...........................................................................................
sql = Aggiornamento__REALESTATE__52(conn)
CALL DB.Execute(sql, 453)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__52(conn)
end if
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 454
'...........................................................................................
sql = Aggiornamento__REALESTATE__53(conn)
CALL DB.Execute(sql, 454)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 455
'...........................................................................................
sql = Aggiornamento__REALESTATE__54(conn)
CALL DB.Execute(sql, 455)
'*******************************************************************************************
'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(455)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 456
'...........................................................................................
sql = Aggiornamento__REALESTATE__55(conn)
CALL DB.Execute(sql, 456)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 457
'...........................................................................................
sql = Aggiornamento__REALESTATE__56(conn)
CALL DB.Execute(sql, 457)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 458
'...........................................................................................
sql = Aggiornamento__REALESTATE__57(conn)
CALL DB.Execute(sql, 458)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 459
'...........................................................................................
sql = Aggiornamento__REALESTATE__58(conn)
CALL DB.Execute(sql, 459)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 460
'...........................................................................................
sql = Aggiornamento__REALESTATE__59(conn)
CALL DB.Execute(sql, 460)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 461
'...........................................................................................
sql = Aggiornamento__REALESTATE__60(conn)
CALL DB.Execute(sql, 461)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 462
'...........................................................................................
sql = Aggiornamento__REALESTATE__61(conn)
CALL DB.Execute(sql, 462)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 463
'...........................................................................................
sql = Aggiornamento__REALESTATE__62(conn)
CALL DB.Execute(sql, 463)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 464
'...........................................................................................
sql = Aggiornamento__REALESTATE__65(conn)
CALL DB.Execute(sql, 464)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 465
'...........................................................................................
sql = Aggiornamento__REALESTATE__66(conn)
CALL DB.Execute(sql, 465)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__66(conn)
end if
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 466
'...........................................................................................
sql = Aggiornamento__REALESTATE__67(conn)
CALL DB.Execute(sql, 466)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 467
'...........................................................................................
sql = Aggiornamento__REALESTATE__68(conn)
CALL DB.Execute(sql, 467)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 468
'...........................................................................................
sql = Aggiornamento__REALESTATE__69(conn)
CALL DB.Execute(sql, 468)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 469
'...........................................................................................
sql = Aggiornamento__REALESTATE__70(conn)
CALL DB.Execute(sql, 469)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 470
'...........................................................................................
sql = Aggiornamento__REALESTATE__71(conn)
CALL DB.Execute(sql, 470)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 471
'...........................................................................................
sql = Aggiornamento__REALESTATE__72(conn)
CALL DB.Execute(sql, 471)
'*******************************************************************************************
'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(471)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 472
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__224(conn)
CALL DB.Execute(sql, 472)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__224(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 473
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__225(conn)
CALL DB.ProtectedExecuteRebuild(sql, 473, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 474
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__226(conn)
CALL DB.ProtectedExecuteRebuild(sql, 474, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 475
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__227(conn)
CALL DB.ProtectedExecuteRebuild(sql, 475, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 476
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__228(conn)
CALL DB.ProtectedExecuteRebuild(sql, 476, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 477
'...........................................................................................
sql = Aggiornamento__REALESTATE__73(conn)
CALL DB.Execute(sql, 477)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__73(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 478
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__229(conn)
CALL DB.ProtectedExecuteRebuild(sql, 478, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 479
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__230(conn)
CALL DB.Execute(sql, 479)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__230(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 480
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__231(conn)
CALL DB.ProtectedExecuteRebuild(sql, 480, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 481
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__232(conn)
CALL DB.ProtectedExecuteRebuild(sql, 481, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 482
'...........................................................................................
sql = Aggiornamento__MEMO2__12(conn)
CALL DB.Execute(sql, 482)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 483
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__233(conn)
CALL DB.ProtectedExecuteRebuild(sql, 483, false, false)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 484
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__234(conn)
CALL DB.ProtectedExecuteRebuild(sql, 484, false, false)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 485
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__235(conn)
CALL DB.ProtectedExecuteRebuild(sql, 485, false, false)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 486
'...........................................................................................
sql = Aggiornamento__REALESTATE__74(conn)
CALL DB.ProtectedExecuteRebuild(sql, 486, false, false)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 487
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__236(conn)
CALL DB.ProtectedExecuteRebuild(sql, 487, false, false)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 488
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__237(conn)
CALL DB.ProtectedExecuteRebuild(sql, 488, false, false)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 489
'...........................................................................................
sql = Aggiornamento__MEMO2__13(conn)
CALL DB.Execute(sql, 489)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 490
'...........................................................................................
if lCase(GetDatabaseName(conn)) = "prestigeinternational" then
	sql = "SELECT * FROM AA_versione"
else
	sql = Aggiornamento__FRAMEWORK_CORE__238(conn)
end if
CALL DB.ProtectedExecuteRebuild(sql, 490, false, false)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 491
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__239(conn)
CALL DB.ProtectedExecuteRebuild(sql, 491, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 492
'...........................................................................................
sql = Aggiornamento__REALESTATE__75(conn)
CALL DB.Execute(sql, 492)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__75(conn)
end if
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 493
'...........................................................................................
sql = Aggiornamento__REALESTATE__76(conn)
CALL DB.ProtectedExecuteRebuild(sql, 493, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 494
'...........................................................................................
sql = Aggiornamento__REALESTATE__77(conn)
CALL DB.ProtectedExecuteRebuild(sql, 494, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 495
'...........................................................................................
sql = Aggiornamento__REALESTATE__78(conn)
CALL DB.ProtectedExecuteRebuild(sql, 495, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 496
'...........................................................................................
sql = Aggiornamento__REALESTATE__79(conn)
CALL DB.Execute(sql, 496)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__79(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 497
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__240(conn)
CALL DB.ProtectedExecuteRebuild(sql, 497, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 498
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__241(conn)
CALL DB.ProtectedExecuteRebuild(sql, 498, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 499
'...........................................................................................
sql = Aggiornamento__REALESTATE__80(conn)
CALL DB.ProtectedExecuteRebuild(sql, 499, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 500
'...........................................................................................
sql = Aggiornamento__MEMO2__14(conn)
CALL DB.Execute(sql, 500)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 501
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__242(conn)
CALL DB.ProtectedExecuteRebuild(sql, 501, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 502
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__243(conn)
CALL DB.ProtectedExecuteRebuild(sql, 502, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 503
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__244(conn)
CALL DB.ProtectedExecuteRebuild(sql, 503, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 504
'...........................................................................................
sql = Aggiornamento__REALESTATE__81(conn)
CALL DB.ProtectedExecuteRebuild(sql, 504, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO PER ALLINEARE PRESTIGE AL MODEL DEL BIPLATFORM ( Nicola - 21/02/2017)
'...........................................................................................
if lCase(GetDatabaseName(conn)) = "prestigeinternational" then
	sql = Install_B2B__version_79()
	CALL DB.Execute(sql, 505)
	
	sql = Indexing_B2B()
	CALL DB.Execute(sql, 506)
	
	sql = Aggiornamento__B2B__1(conn)
	CALL DB.Execute(sql, 507)
	
	sql = Aggiornamento__B2B__2(conn)
	CALL DB.Execute(sql, 508)
	
	sql = Aggiornamento__B2B__3(conn)
	CALL DB.Execute(sql, 509)
	
	sql = Aggiornamento__B2B__4(conn)
	CALL DB.Execute(sql, 510)
	
	sql = Aggiornamento__B2B__5(conn)
	CALL DB.Execute(sql, 511)
	
	sql = Aggiornamento__B2B__6(conn)
	CALL DB.Execute(sql, 512)
	
	sql = Aggiornamento__B2B__7(conn)
	CALL DB.Execute(sql, 513)
	
	sql = Aggiornamento__B2B__8(conn)
	CALL DB.Execute(sql, 514)
	
	sql = Aggiornamento__B2B__9(conn)
	CALL DB.Execute(sql, 515)
	
	sql = Aggiornamento__B2B__10(conn)
	CALL DB.Execute(sql, 516)
	
	sql = Aggiornamento__B2B__11(conn)
	CALL DB.Execute(sql, 517)
	
	sql = Aggiornamento__B2B__12(conn)
	CALL DB.Execute(sql, 518)
	
	sql = Aggiornamento__B2B__13(conn)
	CALL DB.Execute(sql, 519)
	
	sql = Aggiornamento__B2B__14(conn)
	CALL DB.Execute(sql, 520)
	
	sql = Aggiornamento__B2B__15(conn)
	CALL DB.Execute(sql, 521)
	
	sql = Aggiornamento__B2B__16(conn)
	CALL DB.Execute(sql, 522)
	
	sql = Aggiornamento__B2B__17(conn)
	CALL DB.Execute(sql, 523)
	
	sql = Aggiornamento__B2B__18(conn)
	CALL DB.Execute(sql, 524)
	
	sql = Aggiornamento__B2B__19(conn)
	CALL DB.Execute(sql, 525)
	
	sql = Aggiornamento__B2B__20(conn)
	CALL DB.Execute(sql, 526)
	
	sql = Aggiornamento__B2B__21(conn)
	CALL DB.Execute(sql, 527)
	
	sql = Aggiornamento__B2B__22(conn)
	CALL DB.Execute(sql, 528)
	
	sql = Aggiornamento__B2B__23(conn)
	CALL DB.Execute(sql, 529)
	
	sql = Aggiornamento__B2B__24(conn)
	CALL DB.Execute(sql, 530)
	
	sql = Aggiornamento__B2B__25(conn)
	CALL DB.Execute(sql, 531)
	
	sql = Aggiornamento__B2B__26(conn)
	CALL DB.Execute(sql, 532)
	
	sql = Aggiornamento__B2B__27(conn)
	CALL DB.Execute(sql, 533)
	
	sql = Aggiornamento__B2B__28(conn)
	CALL DB.Execute(sql, 534)
	
	sql = Aggiornamento__B2B__29(conn)
	CALL DB.Execute(sql, 535)
	
	sql = Aggiornamento__B2B__30(conn)
	CALL DB.Execute(sql, 536)
	
	sql = Aggiornamento__B2B__31(conn)
	CALL DB.Execute(sql, 537)
	
	sql = Aggiornamento__B2B__32(conn)
	CALL DB.Execute(sql, 538)
	
	sql = Aggiornamento__B2B__33(conn)
	CALL DB.Execute(sql, 539)
	
	sql = Aggiornamento__B2B__34(conn)
	CALL DB.Execute(sql, 540)
	
	sql = Aggiornamento__B2B__35(conn)
	CALL DB.Execute(sql, 541)
	
	sql = Aggiornamento__B2B__36(conn)
	CALL DB.Execute(sql, 542)
	
	sql = Aggiornamento__B2B__37(conn)
	CALL DB.Execute(sql, 543)
	
	sql = Aggiornamento__B2B__38(conn)
	CALL DB.Execute(sql, 544)
	
	sql = Aggiornamento__B2B__39(conn)
	CALL DB.Execute(sql, 545)
	
	sql = Aggiornamento__B2B__40(conn)
	CALL DB.Execute(sql, 546)
	
	sql = Aggiornamento__B2B__41(conn)
	CALL DB.Execute(sql, 547)
	
	sql = Aggiornamento__B2B__42(conn)
	CALL DB.Execute(sql, 548)
	
	sql = Aggiornamento__B2B__43(conn)
	CALL DB.Execute(sql, 549)
	
	sql = Aggiornamento__B2B__44(conn)
	CALL DB.Execute(sql, 550)
	
	sql = Aggiornamento__B2B__45(conn)
	CALL DB.Execute(sql, 551)
	
	sql = Aggiornamento__B2B__46(conn)
	CALL DB.Execute(sql, 552)
	
	sql = Aggiornamento__B2B__47(conn)
	CALL DB.Execute(sql, 553)
	
	sql = Aggiornamento__B2B__48(conn)
	CALL DB.Execute(sql, 554)
	
	sql = Aggiornamento__B2B__49(conn)
	CALL DB.Execute(sql, 555)
	
	sql = Aggiornamento__B2B__50(conn)
	CALL DB.Execute(sql, 556)
	
	sql = Aggiornamento__B2B__51(conn)
	CALL DB.Execute(sql, 557)
	
	sql = Aggiornamento__B2B__52(conn)
	CALL DB.Execute(sql, 558)
	
	sql = Aggiornamento__B2B__53(conn)
	CALL DB.Execute(sql, 559)
	
	sql = Aggiornamento__B2B__54(conn)
	CALL DB.Execute(sql, 560)
	
	sql = Aggiornamento__B2B__55(conn)
	CALL DB.Execute(sql, 561)
	
	sql = Aggiornamento__B2B__56(conn)
	CALL DB.Execute(sql, 562)
	
	sql = Aggiornamento__B2B__57(conn)
	CALL DB.Execute(sql, 563)
	
	sql = Aggiornamento__B2B__58(conn)
	CALL DB.Execute(sql, 564)
	
	sql = Aggiornamento__B2B__59(conn)
	CALL DB.Execute(sql, 565)
	
	sql = Aggiornamento__B2B__60(conn)
	CALL DB.Execute(sql, 566)
	
	sql = Aggiornamento__B2B__61(conn)
	CALL DB.Execute(sql, 567)
	
	sql = Aggiornamento__B2B__62(conn)
	CALL DB.Execute(sql, 568)
	
	sql = Aggiornamento__B2B__63(conn)
	CALL DB.Execute(sql, 569)
	
	sql = Aggiornamento__B2B__64(conn)
	CALL DB.Execute(sql, 570)
	
	sql = Aggiornamento__B2B__65(conn)
	CALL DB.Execute(sql, 571)
	
	sql = Aggiornamento__B2B__66(conn)
	CALL DB.Execute(sql, 572)
	
	sql = Aggiornamento__B2B__67(conn)
	CALL DB.Execute(sql, 573)
	
	sql = Aggiornamento__B2B__68(conn)
	CALL DB.Execute(sql, 574)
	
	sql = Aggiornamento__B2B__69(conn)
	CALL DB.Execute(sql, 575)
	
	sql = Aggiornamento__B2B__70(conn)
	CALL DB.Execute(sql, 576)
	
	sql = Aggiornamento__B2B__71(conn)
	CALL DB.Execute(sql, 577)
	
	sql = Aggiornamento__B2B__72(conn)
	CALL DB.Execute(sql, 578)
	
	sql = Aggiornamento__B2B__73(conn)
	CALL DB.Execute(sql, 579)
	
	sql = Aggiornamento__B2B__74(conn)
	CALL DB.Execute(sql, 580)
	
	sql = Aggiornamento__B2B__75(conn)
	CALL DB.Execute(sql, 581)
	
	sql = Aggiornamento__B2B__76(conn)
	CALL DB.Execute(sql, 582)
	
	sql = Aggiornamento__B2B__77(conn)
	CALL DB.Execute(sql, 583)
	
	sql = Aggiornamento__B2B__78(conn)
	CALL DB.Execute(sql, 584)
	
	sql = Aggiornamento__B2B__79(conn)
	CALL DB.Execute(sql, 585)
	
	sql = Aggiornamento__B2B__80(conn)
	CALL DB.Execute(sql, 586)
	
	sql = Aggiornamento__B2B__81(conn)
	CALL DB.Execute(sql, 587)
	
	sql = Aggiornamento__B2B__82(conn)
	CALL DB.Execute(sql, 588)
	
	sql = Aggiornamento__B2B__83(conn)
	CALL DB.Execute(sql, 589)
	
	sql = Aggiornamento__B2B__84(conn)
	CALL DB.Execute(sql, 590)
	
	sql = Aggiornamento__B2B__85(conn)
	CALL DB.Execute(sql, 591)
	
	sql = Aggiornamento__B2B__86(conn)
	CALL DB.Execute(sql, 592)
	
	sql = Aggiornamento__B2B__87(conn)
	CALL DB.Execute(sql, 593)
	
	sql = Aggiornamento__B2B__88(conn)
	CALL DB.Execute(sql, 594)
	
	sql = Aggiornamento__B2B__89(conn)
	CALL DB.Execute(sql, 595)
	
	sql = Aggiornamento__B2B__90(conn)
	CALL DB.Execute(sql, 596)
	
	sql = Aggiornamento__B2B__91(conn)
	CALL DB.Execute(sql, 597)
	
	sql = Aggiornamento__B2B__92(conn)
	CALL DB.Execute(sql, 598)
	
	sql = Aggiornamento__B2B__93(conn)
	CALL DB.Execute(sql, 599)
	
	sql = Aggiornamento__B2B__94(conn)
	CALL DB.Execute(sql, 600)
	
	sql = Aggiornamento__B2B__95(conn)
	CALL DB.Execute(sql, 601)
	
	sql = Aggiornamento__B2B__96(conn)
	CALL DB.Execute(sql, 602)
	
	sql = Aggiornamento__B2B__97(conn)
	CALL DB.Execute(sql, 603)
	
	sql = Aggiornamento__B2B__98(conn)
	CALL DB.Execute(sql, 604)
	
	sql = Aggiornamento__B2B__99(conn)
	CALL DB.Execute(sql, 605)
	
	sql = Aggiornamento__B2B__100(conn)
	CALL DB.Execute(sql, 606)
	
	sql = Aggiornamento__B2B__101(conn)
	CALL DB.Execute(sql, 607)
	
	sql = Aggiornamento__B2B__102(conn)
	CALL DB.Execute(sql, 608)
	
	sql = Aggiornamento__B2B__103(conn)
	CALL DB.Execute(sql, 609)
	
	sql = Aggiornamento__B2B__104(conn)
	CALL DB.Execute(sql, 610)
	
	sql = Aggiornamento__B2B__105(conn)
	CALL DB.Execute(sql, 611)
	
	sql = Aggiornamento__B2B__106(conn)
	CALL DB.Execute(sql, 612)
	
	sql = Aggiornamento__B2B__107(conn)
	CALL DB.Execute(sql, 613)
	
	sql = Aggiornamento__B2B__108(conn)
	CALL DB.Execute(sql, 614)
	
	sql = Aggiornamento__B2B__109(conn)
	CALL DB.Execute(sql, 615)
	
	sql = Aggiornamento__B2B__110(conn)
	CALL DB.Execute(sql, 616)
	
	sql = Aggiornamento__B2B__111(conn)
	CALL DB.Execute(sql, 617)
	
	sql = Aggiornamento__B2B__112(conn)
	CALL DB.Execute(sql, 618)
	
	sql = Aggiornamento__B2B__113(conn)
	CALL DB.Execute(sql, 619)
	
	sql = Aggiornamento__B2B__114(conn)
	CALL DB.Execute(sql, 620)
	
	sql = Aggiornamento__B2B__115(conn)
	CALL DB.Execute(sql, 621)
	
	sql = Aggiornamento__B2B__116(conn)
	CALL DB.Execute(sql, 622)
	
	sql = Aggiornamento__B2B__117(conn)
	CALL DB.Execute(sql, 623)
	
	sql = Aggiornamento__B2B__118(conn)
	CALL DB.Execute(sql, 624)
	
	sql = Aggiornamento__B2B__119(conn)
	CALL DB.Execute(sql, 625)
	
	sql = Aggiornamento__B2B__120(conn)
	CALL DB.Execute(sql, 626)
	
	sql = Aggiornamento__B2B__121(conn)
	CALL DB.Execute(sql, 627)
	
	sql = Aggiornamento__B2B__122(conn)
	CALL DB.Execute(sql, 628)
	
	sql = Aggiornamento__B2B__123(conn)
	CALL DB.Execute(sql, 629)
	
	sql = Aggiornamento__B2B__124(conn)
	CALL DB.Execute(sql, 630)
	
	sql = Aggiornamento__B2B__125(conn)
	CALL DB.Execute(sql, 631)
	
	sql = Aggiornamento__B2B__126(conn)
	CALL DB.Execute(sql, 632)
	
	sql = Aggiornamento__B2B__127(conn)
	CALL DB.Execute(sql, 633)
	
	sql = Aggiornamento__B2B__128(conn)
	CALL DB.Execute(sql, 634)
	
	sql = Aggiornamento__B2B__129(conn)
	CALL DB.Execute(sql, 635)
	
	sql = Aggiornamento__B2B__130(conn)
	CALL DB.Execute(sql, 636)
	
	sql = Aggiornamento__B2B__131(conn)
	CALL DB.Execute(sql, 637)
	
	sql = Aggiornamento__B2B__132(conn)
	CALL DB.Execute(sql, 638)
	
	sql = Aggiornamento__B2B__133(conn)
	CALL DB.Execute(sql, 639)
	
	sql = Aggiornamento__B2B__134(conn)
	CALL DB.Execute(sql, 640)
	
	sql = Aggiornamento__B2B__135(conn)
	CALL DB.Execute(sql, 641)
	
	sql = Aggiornamento__B2B__136(conn)
	CALL DB.Execute(sql, 642)
	
	sql = Aggiornamento__B2B__137(conn)
	CALL DB.Execute(sql, 643)
	
	sql = Aggiornamento__B2B__138(conn)
	CALL DB.Execute(sql, 644)
	
	sql = Aggiornamento__B2B__139(conn)
	CALL DB.Execute(sql, 645)
	
	sql = Aggiornamento__B2B__140(conn)
	CALL DB.Execute(sql, 646)
	
	sql = Aggiornamento__B2B__141(conn)
	CALL DB.Execute(sql, 647)
	
	sql = Aggiornamento__B2B__142(conn)
	CALL DB.Execute(sql, 648)
	
	sql = Aggiornamento__B2B__143(conn)
	CALL DB.Execute(sql, 649)
	
	sql = Aggiornamento__B2B__144(conn)
	CALL DB.Execute(sql, 650)
	
	sql = Aggiornamento__B2B__145(conn)
	CALL DB.Execute(sql, 651)
	
	sql = Aggiornamento__B2B__146(conn)
	CALL DB.Execute(sql, 652)
	
	sql = Aggiornamento__B2B__147(conn)
	CALL DB.Execute(sql, 653)
	
	sql = Aggiornamento__B2B__148(conn)
	CALL DB.Execute(sql, 654)
	
	sql = Aggiornamento__B2B__149(conn)
	CALL DB.Execute(sql, 655)
	
	sql = Aggiornamento__B2B__150(conn)
	CALL DB.Execute(sql, 656)
	
	sql = Aggiornamento__B2B__151(conn)
	CALL DB.Execute(sql, 657)
	
	sql = Aggiornamento__B2B__152(conn)
	CALL DB.Execute(sql, 658)
	
	sql = Aggiornamento__B2B__153(conn)
	CALL DB.Execute(sql, 659)
	
	sql = Aggiornamento__B2B__154(conn)
	CALL DB.Execute(sql, 660)
	
	sql = Aggiornamento__B2B__155(conn)
	CALL DB.Execute(sql, 661)
	
	sql = Aggiornamento__B2B__156(conn)
	CALL DB.Execute(sql, 662)
	
	sql = Aggiornamento__B2B__157(conn)
	CALL DB.Execute(sql, 663)
	
	sql = Aggiornamento__B2B__158(conn)
	CALL DB.Execute(sql, 664)
	
	sql = Aggiornamento__B2B__159(conn)
	CALL DB.Execute(sql, 665)
	
	sql = Aggiornamento__B2B__160(conn)
	CALL DB.Execute(sql, 666)
	
	sql = Aggiornamento__B2B__161(conn)
	CALL DB.Execute(sql, 667)
	
	sql = Aggiornamento__B2B__162(conn)
	CALL DB.Execute(sql, 668)
	
	sql = Aggiornamento__B2B__163(conn)
	CALL DB.Execute(sql, 669)
	
	sql = Aggiornamento__B2B__164(conn)
	CALL DB.Execute(sql, 670)
	
	sql = Aggiornamento__B2B__165(conn)
	CALL DB.Execute(sql, 671)
	
	sql = Aggiornamento__B2B__166(conn)
	CALL DB.Execute(sql, 672)
	
	sql = Aggiornamento__B2B__167(conn)
	CALL DB.Execute(sql, 673)
	
	sql = Aggiornamento__B2B__168(conn)
	CALL DB.Execute(sql, 674)
	
	sql = Aggiornamento__B2B__169(conn)
	CALL DB.Execute(sql, 675)
	
	sql = Aggiornamento__B2B__170(conn)
	CALL DB.Execute(sql, 676)
	
	sql = Aggiornamento__B2B__171(conn)
	CALL DB.Execute(sql, 677)
	
	sql = Aggiornamento__B2B__172(conn)
	CALL DB.Execute(sql, 678)
	
	sql = Aggiornamento__B2B__173(conn)
	CALL DB.Execute(sql, 679)
	
	sql = Aggiornamento__B2B__174(conn)
	CALL DB.Execute(sql, 680)
	
	sql = Aggiornamento__B2B__175(conn)
	CALL DB.Execute(sql, 681)
	
	sql = Aggiornamento__B2B__176(conn)
	CALL DB.Execute(sql, 682)
	
	sql = Aggiornamento__B2B__177(conn)
	CALL DB.Execute(sql, 683)
	
	sql = Aggiornamento__B2B__178(conn)
	CALL DB.Execute(sql, 684)
	
	sql = Aggiornamento__B2B__179(conn)
	CALL DB.Execute(sql, 685)
	
	sql = Aggiornamento__B2B__180(conn)
	CALL DB.Execute(sql, 686)
	
	sql = Aggiornamento__B2B__181(conn)
	CALL DB.Execute(sql, 687)
	
	sql = Aggiornamento__B2B__182(conn)
	CALL DB.Execute(sql, 688)
	
	sql = Aggiornamento__B2B__183(conn)
	CALL DB.Execute(sql, 689)
	
	sql = Aggiornamento__B2B__184(conn)
	CALL DB.Execute(sql, 690)
	
	sql = Aggiornamento__B2B__185(conn)
	CALL DB.Execute(sql, 691)
	
	sql = Aggiornamento__B2B__186(conn)
	CALL DB.Execute(sql, 692)
	
	sql = Aggiornamento__B2B__187(conn)
	CALL DB.Execute(sql, 693)
	
	sql = Aggiornamento__B2B__188(conn)
	CALL DB.Execute(sql, 694)
	
	sql = Aggiornamento__B2B__189(conn)
	CALL DB.Execute(sql, 695)
	
	sql = Aggiornamento__B2B__190(conn)
	CALL DB.Execute(sql, 696)
	
	sql = Aggiornamento__B2B__191(conn)
	CALL DB.Execute(sql, 697)
	
	sql = Aggiornamento__B2B__192(conn)
	CALL DB.Execute(sql, 698)
	
	sql = Aggiornamento__B2B__193(conn)
	CALL DB.Execute(sql, 699)
	
	sql = Aggiornamento__B2B__194(conn)
	CALL DB.Execute(sql, 700)
	
	sql = Aggiornamento__B2B__195(conn)
	CALL DB.Execute(sql, 701)
	
	sql = Aggiornamento__B2B__196(conn)
	CALL DB.Execute(sql, 702)
	
	sql = Aggiornamento__B2B__197(conn)
	CALL DB.Execute(sql, 703)
	
	sql = Aggiornamento__B2B__198(conn)
	CALL DB.Execute(sql, 704)
	
	sql = Aggiornamento__B2B__199(conn)
	CALL DB.Execute(sql, 705)
	
	sql = Aggiornamento__B2B__200(conn)
	CALL DB.Execute(sql, 706)
	
	sql = Aggiornamento__B2B__201(conn)
	CALL DB.Execute(sql, 707)
	
	sql = Aggiornamento__B2B__202(conn)
	CALL DB.Execute(sql, 708)
	
	sql = Aggiornamento__B2B__203(conn)
	CALL DB.Execute(sql, 709)
	
	sql = Aggiornamento__B2B__204(conn)
	CALL DB.Execute(sql, 710)
	
	sql = Aggiornamento__B2B__205(conn)
	CALL DB.Execute(sql, 711)
	
	sql = Aggiornamento__B2B__206(conn)
	CALL DB.Execute(sql, 712)
	
	sql = Aggiornamento__B2B__207(conn)
	CALL DB.Execute(sql, 713)
	
	sql = Aggiornamento__B2B__208(conn)
	CALL DB.Execute(sql, 714)
	
	sql = Aggiornamento__B2B__209(conn)
	CALL DB.Execute(sql, 715)
	
	sql = Aggiornamento__B2B__210(conn)
	CALL DB.Execute(sql, 716)
	
	sql = Aggiornamento__B2B__211(conn)
	CALL DB.Execute(sql, 717)
	
	sql = Aggiornamento__B2B__212(conn)
	CALL DB.Execute(sql, 718)
	
	sql = Aggiornamento__B2B__213(conn)
	CALL DB.Execute(sql, 719)
	
	sql = Aggiornamento__B2B__214(conn)
	CALL DB.Execute(sql, 720)
	
	sql = Aggiornamento__B2B__215(conn)
	CALL DB.Execute(sql, 721)
	
	sql = Aggiornamento__B2B__216(conn)
	CALL DB.Execute(sql, 722)
	
	sql = Aggiornamento__B2B__217(conn)
	CALL DB.Execute(sql, 723)
	
	sql = Aggiornamento__B2B__218(conn)
	CALL DB.Execute(sql, 724)
	
	sql = Aggiornamento__B2B__219(conn)
	CALL DB.Execute(sql, 725)
	
	sql = Aggiornamento__B2B__220(conn)
	CALL DB.Execute(sql, 726)
	
	sql = Aggiornamento__B2B__221(conn)
	CALL DB.Execute(sql, 727)
	
	sql = Aggiornamento__B2B__222(conn)
	CALL DB.Execute(sql, 728)
	
	sql = Aggiornamento__B2B__223(conn)
	CALL DB.Execute(sql, 729)
	
	sql = Aggiornamento__B2B__224(conn)
	CALL DB.Execute(sql, 730)
	
	sql = Aggiornamento__B2B__225(conn)
	CALL DB.Execute(sql, 731)
	
	sql = Aggiornamento__B2B__226(conn)
	CALL DB.Execute(sql, 732)
	
	sql = Aggiornamento__B2B__227(conn)
	CALL DB.Execute(sql, 733)
	
	sql = "UPDATE AA_versione SET versione=504"
	CALL DB.Execute(sql, 734)
end if
'*******************************************************************************************

'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************

'PASSAGGIO A BIPLATFORM

'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'AGGIORNAMENTO 505
'...........................................................................................
sql = Aggiornamento__BiPlatform__00000(conn)
CALL DB.Execute(sql, 505)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 506
'...........................................................................................
sql = Aggiornamento__BiPlatform__00001(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 507
'...........................................................................................
sql = Aggiornamento__BiPlatform__00002(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 508
'...........................................................................................
sql = Aggiornamento__BiPlatform__00003(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 509
'...........................................................................................
sql = Aggiornamento__BiPlatform__00004(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 510
'...........................................................................................
sql = Aggiornamento__BiPlatform__00005(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 511
'...........................................................................................
sql = Aggiornamento__BiPlatform__00006(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 512
'...........................................................................................
sql = Aggiornamento__BiPlatform__00007(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 513
'...........................................................................................
sql = Aggiornamento__BiPlatform__00008(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 514
'...........................................................................................
sql = Aggiornamento__BiPlatform__00009(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 515
'...........................................................................................
sql = Aggiornamento__BiPlatform__00010(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 516
'...........................................................................................
sql = Aggiornamento__BiPlatform__00011(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 517
'...........................................................................................
sql = Aggiornamento__BiPlatform__00012(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 518
'...........................................................................................
sql = Aggiornamento__BiPlatform__00013(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 519
'...........................................................................................
sql = Aggiornamento__BiPlatform__00014(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 520
'...........................................................................................
sql = Aggiornamento__BiPlatform__00015(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 521
'...........................................................................................
sql = Aggiornamento__BiPlatform__00016(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 522
'...........................................................................................
sql = Aggiornamento__BiPlatform__00017(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 523
'...........................................................................................
sql = Aggiornamento__BiPlatform__00018(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 524
'...........................................................................................
sql = Aggiornamento__BiPlatform__00019(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 525
'...........................................................................................
sql = Aggiornamento__BiPlatform__00020(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 526
'...........................................................................................
sql = Aggiornamento__BiPlatform__00021(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 527
'...........................................................................................
sql = Aggiornamento__BiPlatform__00022(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 528
'...........................................................................................
sql = Aggiornamento__BiPlatform__00023(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 529
'...........................................................................................
sql = Aggiornamento__BiPlatform__00024(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 530
'...........................................................................................
sql = Aggiornamento__BiPlatform__00025(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 531
'...........................................................................................
sql = Aggiornamento__BiPlatform__00026(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 532
'...........................................................................................
sql = Aggiornamento__BiPlatform__00027(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 533
'...........................................................................................
sql = Aggiornamento__BiPlatform__00028(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 534
'...........................................................................................
sql = Aggiornamento__BiPlatform__00029(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 535
'...........................................................................................
sql = Aggiornamento__BiPlatform__00030(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 536
'...........................................................................................
sql = Aggiornamento__BiPlatform__00031(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 537
'...........................................................................................
sql = Aggiornamento__BiPlatform__00032(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 538
'...........................................................................................
sql = Aggiornamento__BiPlatform__00033(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 539
'...........................................................................................
sql = Aggiornamento__BiPlatform__00034(conn)
CALL DB.SequentialExecute(sql)
'*******************************************************************************************

%>
<% '........................................................................................... %>
<!--#INCLUDE FILE="Update__FileFooter.asp" -->
<% '........................................................................................... %>
