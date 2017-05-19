DECLARE Sleep IN Win32API INTEGER
EF = CREATEOBJECT("Excel.Application")
EF.Visible = .T.
EF.Workbooks.Add
DO WHILE ISNULL(EF.Cells(1, 1).Value)
		=Sleep(1000)
ENDDO 
DO WHILE EF.Visible
	i = 1
	CREATE CURSOR sh_vrem (nom c(25) null, ei c(100) null) 
	DO WHILE .NOT. ISNULL(EF.Cells(i,1).value)
		IF (TYPE("EF.Cells(i,1).value")#"N" AND TYPE("EF.Cells(i,1).value")#"D";
	    AND LEN(ALLTRIM(EF.Cells(i,1).value))#0) OR TYPE("EF.Cells(i,1).value")="N";
	    OR TYPE("EF.Cells(i,1).value")= "D"
			mnom=EF.Cells(i,1).value  
			IF TYPE("mnom")="N"  
				mnom=ALLTRIM(STR(mnom))  
			ELSE 
				IF TYPE("mnom")="D"  
					mnom=ALLTRIM(dtoc(mnom))  
				ELSE 
					mnom=ALLTRIM(mnom)  
				ENDIF 
			ENDIF 
			mei=ALLTRIM(EF.Cells(i,2).value)  
			IF TYPE("mei")="N"  
				mei=ALLTRIM(STR(mei)) 
			ELSE    	  
				IF TYPE("mei")="D"  
					mei=ALLTRIM(dtoc(mei))    
				ENDIF
			ENDIF  
			SELECT sh_vrem   
			APPEND BLANK  
			REPLACE nom WITH mnom, ei WITH mei
		ENDIF 
		i = i + 1
	ENDDO   
	=Sleep(1000)
ENDDO 
SELECT sh_vrem
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
ELSE
	SQLEXEC(con_bd1, "select nn,ei,nm,marka,razm,cd,cs,cm,obc,cenapost,dokosn,gsn,gm,post,inn,adv,id_order from uit.dbo.prnom_p where usersh=?fam order by id_order",'prnom_prosm')
ENDIF
temp_baz='prnom_prosm'
DO FORM prnom_pr