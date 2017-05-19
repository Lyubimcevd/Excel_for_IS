PARAMETERS vx_pshifr
	SELECT 0
	FileName=GETFILE('xls;xlsx;dbf','','Выбрать',0,'Открыть файл ном.номеров...')
	IF !EMPTY(FileName) and UPPER(RIGHT(RTRIM(FileName),3))='DBF'&&FILE('c:\tru\prnome.dbf')
		USE &FileName
		GO top
		IF RECCOUNT()#0
			USE
			CREATE cursor sh_vrem (nom c(25),ei c(2))
			APPEND FROM &FileName
		ELSE
			return
		endif
	ELSE
		IF 'XLS'$UPPER(RIGHT(RTRIM(FileName),4))
			CREATE CURSOR sh_vrem (nom c(25) null, ei c(2) null)
			XLRelek1=GetObject("","Excel.Application")  
			XLRelek=XLRelek1.Workbooks.open((FileName))  
*!*				XLRelek.application.visible=.t.
			XLSheet1=XLRelek.Sheets(1)  &&Первый лист
			XLSheet1.select   
			ii=1 &&
			XLSheet1.Cells(ii,1).select   
			DO WHILE IIF(TYPE("XLSheet1.Cells(ii,1).value")="N",STR(XLSheet1.Cells(ii,1).value),IIF(TYPE("XLSheet1.Cells(ii,1).value")="D",DTOC(XLSheet1.Cells(ii,1).value),XLSheet1.Cells(ii,1).value))#SPACE(25)&&!EMPTY(XLSheet1.Cells(ii,1).value)
				STORE '' TO mnom,mei
			   	XLSheet1.Cells(ii,1).select &&[1]  
			  	mnom=XLSheet1.Cells(ii,1).value  
			  	IF TYPE("mnom")="N"  
			  		mnom=rTRIM(STR(mnom))  
			  	ENDIF   
			  	IF TYPE("mnom")="D"  
			  		mnom=rTRIM(dtoc(mnom))  
			  	ENDIF   
			  	mnom=rTRIM(mnom)  
			  	IF ISNULL(mnom)=.t.
			  		mnom=''
			  	endif  	  
			  	XLSheet1.Cells(ii,2).select &&[1]  
			  	mei=ALLTRIM(XLSheet1.Cells(ii,2).value)  
			  	IF TYPE("mei")="N"  
			  		mei=ALLTRIM(STR(mei))  
			  	ENDIF   	  
			  	IF TYPE("mei")="D"  
			  		mei=ALLTRIM(dtoc(mei))  
			  	ENDIF      
			  	IF ISNULL(mei)=.t.
			  		mei=''
			  	endif
			  	SELECT sh_vrem   
			  	APPEND BLANK  
				replace nom WITH mnom, ei WITH mei
			  	ii=ii+1  
			*!*	  	XLSheet1.Cells(ii,1).select
			enddo   
			XLRelek.application.quit &&Закрываю. Можете и не закрывать для сверки  
*!*				use
*!*				CREATE cursor sh_vrem (nom c(25),ei c(2))
*!*				APPEND FROM &FileName TYPE XL5
		else
*!*			MESSAGEBOX('Не выбран файл')
			RETURN
		endif
	ENDIF
SQLEXEC(con_bd1, 'delete from uit.dbo.prnom where usersh=?fam','prov')
SET DATE GERMAN
dd=DATE()
IF vx_pshifr#'ei'
	m.ei='  '
endif
SCAN
	SCATTER MEMVAR
	SQLSETPROP(con_bd1, 'Asynchronous', .F.)
	SQLEXEC(con_bd1,"INSERT INTO uit.dbo.prnom (nom,ei,usersh,dvv) values (?m.nom,?m.ei,?fam,?dd)")	
ENDSCAN
USE
IF vx_pshifr='ei'
	SQLEXEC(con_bd1, "select nn,ei,nm,marka,razm,cd,cs,cm,obc,cenapost,dokosn,gsn,gm,post,inn,adv,id_order from uit.dbo.prnom_pei where usersh=?fam order by id_order",'prnom_prosm')
	temp_baz='prnom_prosm'
ELSE
	SQLEXEC(con_bd1, "select nn,ei,nm,marka,razm,cd,cs,cm,obc,cenapost,dokosn,gsn,gm,post,inn,adv,id_order from uit.dbo.prnom_p where usersh=?fam order by id_order",'prnom_prosm')
	temp_baz='prnom_prosm'
ENDIF
DO FORM prnom_pr
&&SELECT nn,ei,nm,marka,razm,gm,gs,obc,prbei,pmc,gsn,cs,cm,dmc_order,cd,dk_order,ck,pk,dk,dmc,cenapost,dokosn
&&,post,inn,adv