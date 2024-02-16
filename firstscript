on processdatenbankcheck(scriptIcon, filePath)
	
	openNumbersFile(filePath)
	
	-- Eingabeaufforderung für Nachnamen und Vornamen
	set lastName to text returned of (display dialog "Bitte gib den Familiennamen ein:" default answer "")
	
	
	-- Anzeigen einer Nachricht während des Wartens
	display dialog "Die Suche wird im Hintergrund ausgeführt. Dies kann einen Moment dauern." & return & return & "Bitte warten Sie bis sich das nächste Programmfenster mit weiteren Anweisungen öffnet..." buttons {} giving up after 3
	
	-- Numbers-Dokument öffnen und Suche durchführen
	tell application "Numbers"
		activate
		set myDocument to front document
		set targetSheet to sheet 7 of myDocument
		set targetTable to table 1 of targetSheet
		set lastNameColumn to column "G" of targetTable
		set firstNameColumn to column "H" of targetTable
		
		-- Liste zum Speichern der gefundenen Datensätze
		set foundRecords to {}
		set foundRows to {}
		
		-- Durchsuche die Tabelle nach dem Familiennamen und Namen
		repeat with i from 1 to count of rows of targetTable
			set currentLastName to value of cell i of lastNameColumn as string
			
			
			if currentLastName is lastName then
				-- Datensatz gefunden, füge die Zeile zur Liste hinzu
				set end of foundRecords to {rowIndex:i as string, customerfirma:(value of cell 3 of row i of targetTable as string), customerNumber:(value of cell 6 of row i of targetTable as string), customer:(value of cell 7 of row i of targetTable as string), Customername:(value of cell 8 of row i of targetTable as string), Customerstr:(value of cell 9 of row i of targetTable as string), Customerplz:(value of cell 10 of row i of targetTable as string), Customerort:(value of cell 11 of row i of targetTable as string)}
				
				set end of foundRows to {rowIndex:i as string}
				
			end if
		end repeat
		
		-- Wenn Datensätze gefunden wurden
		if (count of foundRecords) > 0 then
			-- Liste der gefundenen Datensätze erstellen
			set recordList to {}
			set rowlist to {}
			repeat with foundRecord in foundRecords
				
				set end of recordList to (rowIndex of foundRecord) & " - " & (customerfirma of foundRecord) & " - " & (customerNumber of foundRecord) & " - " & (customer of foundRecord) & ", " & (Customername of foundRecord) & ", " & (Customerstr of foundRecord) & ", " & (Customerplz of foundRecord) & ", " & (Customerort of foundRecord)
				set chosenrowindex to (rowIndex of foundRecord)
				
			end repeat
			
			set end of recordList to "Neuen Kunden anlegen..."
			-- Dialogfeld mit den gefundenen Daten und der Option zum Erstellen eines neuen Datensatzes anzeigen
			set chosenRecord to choose from list recordList with prompt "Mehrere Datensätze gefunden. Bitte wähle einen aus oder erstelle einen neuen Datensatz:" default items {item 1 of recordList}
			
			
			
			-- Hier kannst du den ausgewählten Datensatz (chosenRecord) und den Zeilenindex (chosenRowIndex) weiterverarbeiten			
			-- Behandle den ausgewählten Datensatz oder die Auswahl zum Erstellen eines neuen Datensatzes
			
			if chosenRecord contains "Neuen Kunden anlegen..." then
				-- Code zum Erstellen eines neuen Datensatzes hier einfügen
				display dialog "Neuen Kunden anlegen..." buttons {"OK"} default button "OK"
				
				processneukunde(scriptIcon, filePath)
			end if
			
		else
			-- Der Benutzer hat einen existierenden Datensatz ausgewählt, fahre mit der weiteren Verarbeitung fort
			set chosenrowindex to (word 1 of (item 1 of chosenRecord)) as integer
			display dialog "Du hast den nachstehenden Datensatz ausgewählt, dieser wird in die Buchungsmaske automatisch übernommen " & chosenRecord buttons {"OK"} default button "OK"
			
			-- Hier kannst du chosenRowIndex weiter verwenden
			set chosenrecordbutton to button returned of result
			if chosenrecordbutton is "OK" then
				-- Verarbeite den ausgewählten Datensatz weiter
				
				
				set i to chosenrowindex
				
				set customerpreisstufe to value of cell 15 of row i of targetTable as string
				set customerNumber to value of cell 6 of row i of targetTable as string
				set customerfirma to value of cell 3 of row i of targetTable as string
				set customertelefon to value of cell 12 of row i of targetTable as string
				set customermobile to value of cell 13 of row i of targetTable as string
				set customermail to value of cell 14 of row i of targetTable as string
				set customer to value of cell 7 of row i of targetTable as string
				set Customername to value of cell 8 of row i of targetTable as string
				set Customerstr to value of cell 9 of row i of targetTable as string
				set Customerplz to value of cell 10 of row i of targetTable as string
				set Customerort to value of cell 11 of row i of targetTable as string
				
				-- Trage die Informationen in die entsprechenden Zellen der Tabelle ein
				set value of cell "C4" of table 1 of sheet "Eingabemaske" of myDocument to customerpreisstufe
				set value of cell "C7" of table 1 of sheet "Eingabemaske" of myDocument to customerNumber
				set value of cell "C12" of table 1 of sheet "Eingabemaske" of myDocument to customerfirma
				set value of cell "C13" of table 1 of sheet "Eingabemaske" of myDocument to customer
				set value of cell "C14" of table 1 of sheet "Eingabemaske" of myDocument to Customername
				set value of cell "C15" of table 1 of sheet "Eingabemaske" of myDocument to Customerstr
				set value of cell "C16" of table 1 of sheet "Eingabemaske" of myDocument to Customerplz
				set value of cell "C17" of table 1 of sheet "Eingabemaske" of myDocument to Customerort
				set value of cell "C19" of table 1 of sheet "Eingabemaske" of myDocument to customertelefon
				set value of cell "C20" of table 1 of sheet "Eingabemaske" of myDocument to customermobile
				set value of cell "C21" of table 1 of sheet "Eingabemaske" of myDocument to customermail
				
				-- Bestätigung anzeigen
				set datenuebernahmeDialog to display dialog "Die Daten wurden erfolgreich in die Tabelle eingetragen." buttons {"OK"} default button "OK"
				
				
				
				
				
			end if
			
		end if
		
		-- Wenn kein Datensatz gefunden wurde
		set chosenRecord to choose from list {"Neuen Datensatz erstellen..."} with prompt "Kein Datensatz mit dem eingegebenen Familiennamen gefunden. Möchtest du einen neuen Datensatz erstellen?" default items {"Neuen Datensatz erstellen..."}
		if chosenRecord is not false then
			-- Code zum Erstellen eines neuen Datensatzes hier einfügen
			display dialog "Neuen Datensatz erstellen..." buttons {"OK"} default button "OK"
			
			processneukunde(scriptIcon, filePath)
			clearInputMask(scriptIcon, filePath)
			showMainDialog(scriptIcon, filePath)
		end if
		
		
		
		
		
		
		
		processneukunde(scriptIcon, filePath)
		
		
		
	end tell
	
	
end processdatenbankcheck


on processneukunde(scriptIcon, filePath)
	
	tell application "Numbers"
		-- Aktives Dokument
		set myDocument to front document
		
		-- Quellzelle (hier: Tabellenblatt 1, Spalte B, Zeile 2)
		set sourceCell to cell "B5" of table 1 of sheet "Hilfsformel" of myDocument
		set sourceValue to value of sourceCell
		
		-- Zielzelle (hier: Tabellenblatt 2, Spalte C, Zeile 3)
		set targetCell to cell "C7" of table 1 of sheet "Eingabemaske" of myDocument
		
		-- Kopiere den Wert von der Quellzelle zur Zielzelle
		set value of targetCell to sourceValue
		
	end tell
	
	set inputDataDialog to display dialog "Bitte gib die erforderlichen Daten in der Numbers Datei in die Eingabemaske für die Anzahlungsrechnung ein." & return & return & "Sind alle Daten korrekt eingegeben?" & return & return & return & return & "Bestätige deine Eingabe hier mit 'OK'." & return & return & "Du kannst später keine Korrekturen mehr an den Daten vornehmen!" buttons {"OK", "Hauptmenü"} default button "OK" with icon file scriptIcon with title "ERP Ferienhaus Spreeblick - Datenerfassung"
	
	set inputDataButton to button returned of inputDataDialog
	
	if inputDataButton is "OK" then
		tell application "Numbers"
			activate
			set myDocument to front document
			
			tell document 1
				set active sheet to sheet "Anzahlungsrechnung"
			end tell
			
		end tell
		
		
		set abbrechDialog to display dialog "Schau dir deinen Rechnungsentwurf an, ob alle Daten korrekt übernommen wurden." & return & return & "Ab jetzt beginnt die Rechnungsstellung." & return & return & return & return & "Bestätige deine Eingabe hier mit 'OK'." & return & return & "Du kannst später keine Korrekturen mehr an der Rechnung vornehmen!" buttons {"OK", "Zurück", "Hauptmenü"} default button "OK" with icon file scriptIcon with title "ERP Ferienhaus Spreeblick - Datenerfassung"
		
		
		
		set abbrechButton to button returned of abbrechDialog
		if abbrechButton is "Zurück" then
			-- Dialog für Bestätigung
			set confirmDialog to display dialog "Bist du sicher, dass du die Rechnungsstellung abbrechen möchtest? Du gelangst dann zurück zur Eingabemaske." buttons {"Ja", "Nein"} default button "Nein" with icon file scriptIcon with title "ERP Ferienhaus Spreeblick - Abbrechen"
			
			-- Überprüfen der Bestätigung
			if button returned of confirmDialog is "Ja" then
				-- Aktionen für 'Ja' hier einfügen
				processAnzahlungsrechnung(scriptIcon, filePath)
			else
				-- Fortsetzen mit weiteren Aktionen oder Skript beenden
				processAnzahlungsrechnungneukundeErstellen(scriptIcon, filePath)
			end if
		else if abbrechButton is "OK" then
			-- Weiter mit der Verarbeitung
			processAnzahlungsrechnungneukundeErstellen(scriptIcon, filePath)
		else if abbrechButton is "Hauptmenü" then
			clearInputMask(scriptIcon, filePath)
			showMainDialog(scriptIcon, filePath)
		end if
	else if inputDataButton is "Hauptmenü" then
		clearInputMask(scriptIcon, filePath)
		showMainDialog(scriptIcon, filePath)
	end if
	
end processneukunde


