/* -----------------------------------------------------------------------------------

	ORIGO DATA IMPORTERING

	
	Syntax: mr [kommando] [option] [option] [option] [option] [...]
	
	
----------------------------------------------------------------------------------- */

program drop _all 
discard

program oi 

	if "`1'" == "" {
		di "---------------------------------------------------------------------------------------"
		di "  ORIGO DATA IMPORTERING"
		di "---------------------------------------------------------------------------------------" _newline
		di "    Laddar data från en Excelfil med data i Origos format. Programmet läser in"
		di "    etiketter för variabeler samt värden och applicerar dessa." _newline
		di " 	Format: oi [ filnamn ]" _newline
		di "---------------------------------------------------------------------------------------" 
		exit
		}
	else {
		quietly {
			local fn = "`1'"

			/*
				Ladda in variabelnamn och sparade dem temporärt (i lokalt macro)
			*/	
			
			noisily display "Öppnar fil..."
			capture import excel "`fn'", sheet("Variabler") cellrange(A3) clear allstring
			error _rc
			
			noisily display "Läser in etiketter för variabler..."
			drop if B==""
			gen str kmb = A + "|" + C
			levelsof kmb, local(kombinamn)

			/*
				Ladda in svarsalternativ och sparade dem temporärt (i en do-fil)
			*/	
			noisily display "Läser in etiketter för värden..."
			capture import excel "`fn'", sheet("Svarsalternativ") cellrange(A3) clear
			error _rc

			destring B, generate(B2) force // konverterar variabeln till siffror, missing om värdet är text
			drop if B2 == . // ta bort de labels som hade text som value
			tostring B2, generate(B3) // konverterar tillbaka till strängvariabel

			count
			forval i = 1/`r(N)' {
				if A[`i'] != "" local var = A[`i']
				local val = B3[`i']
				local lbl = C[`i']
				label define lbl_`var' `val' "`lbl'", add
				}

			tempfile labelfile
			label save using `labelfile'

			/*
				Ladda in data och applicera variabelnamn (som variable-labels)
				och svarsalternetiv (som value-labels)
			*/	
			noisily display "Läser in data..."
			capture import excel "`fn'", sheet("Data") firstrow clear
			error _rc

			// Lägger till variabel-labels
			foreach l of local kombinamn { 
				local varname=substr(substr("`l'", 1, strpos("`l'","|")-1),1,32)
				local labelname=substr("`l'", strpos("`l'","|")+1,.)

				capture confirm variable `varname'
				if !_rc { 
					label variable `varname' "`labelname'"
					}
				}

			// Lägger till value-labels
			run "`labelfile'"
			label dir

			foreach l in `r(names)' {
				local varname=substr("`l'", 5,.)
				label values `varname' `l'
			}
			
			noisily display "Klar."
		}
				
	}

end
