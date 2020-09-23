(function () {
	
	var doc = app.documents[0];
	var o_style = app.documents[0].objectStyles.item('sidenote');

	function renumber_notes ()
	{
		var mode = renumber_mode (doc);
		reset_numbers ();
		switch (mode)
		{
			case 'None': break;
			case 'Page': renumber_each_page (); break;
			case 'Spread': renumber_each_spread (); break;
			case 'Section': renumber_each_section (); break;
		}
		doc.crossReferenceSources.everyItem().update();
	}


	function renumber_each_section () {
		var notes = findNotes ();
		var sectionIndex = 0;
		for (var i = 0; i < notes.length; i++){
			if (notes[i].parent.parentTextFrames[0].parentPage.appliedSection.index > sectionIndex){
				notes[i].paragraphs[0].numberingContinue = false;
				sectionIndex = notes[i].parent.parentTextFrames[0].parentPage.appliedSection.index;
			}
		}
	}


	function renumber_each_spread () {
		var notes = findNotes ();
		var spreadIndex = 0;
		for (var i = 0; i < notes.length; i++){
			if (notes[i].parent.parentTextFrames[0].parentPage.parent.index > spreadIndex){
				notes[i].paragraphs[0].numberingContinue = false;
				spreadIndex++;
			}
		}
	}


	function renumber_each_page () {
		var notes = findNotes ();
		var pageIndex = 0;
		for (var i = 0; i < notes.length; i++){
			if (notes[i].parent.parentTextFrames[0].parentPage.documentOffset > pageIndex){
				notes[i].paragraphs[0].numberingContinue = false;
				pageIndex++
			}
		}
	}


	function findNotes () {
		app.findObjectPreferences = null;
		app.findObjectPreferences.appliedObjectStyles = o_style;
		return doc.findObject();
	}


	function reset_numbers () {
		var notes = findNotes ();
		for (var i = 0; i < notes.length; i++) {
			notes[i].paragraphs[0].numberingContinue = true;
		}
	}


	function renumber_mode () {
		var w = new Window ('dialog', "Sidenote options", undefined, {closeButton: false});
			w.alignChildren = "right";
			w.panel = w.add ('panel {orientation: "row"}');
				w.prompt = w.panel.add ('checkbox {text: "Restart numbering every"}');
				w.list = w.panel.add ('dropdownlist', undefined, ['Page', 'Spread', 'Section']);
					w.list.preferredSize.width = 100;
			w.buttons = w.add ('group');
				w.buttons.add ('button', undefined, 'OK', {name: 'ok'});
				w.buttons.add ('button', undefined, 'Cancel', {name: 'cancel'});

		var sidenote_restarting = doc.extractLabel ('sidenote_restarting');
		if (sidenote_restarting != "") {
			w.prompt.value = true;
			switch (sidenote_restarting) {
				case 'None': w.list.selection = 0; w.list.enabled = false; w.prompt.value = false; break;
				case 'Page': w.list.selection = 0; break;
				case 'Spread': w.list.selection = 1; break;
				case 'Section': w.list.selection = 2; break;
			}
		} else {
			w.list.selection = 0;
			w.list.enabled = false;
		}
		
		w.prompt.onClick = function () {w.list.enabled = this.value}

		if (w.show() == 1) {
			if (w.prompt.value) {
				doc.insertLabel ('sidenote_restarting', w.list.selection.text);
				return w.list.selection.text;
			} else {
				doc.insertLabel ('sidenote_restarting', 'None');
				return 'None';
			}
		}
		exit();
	}// renumber_mode


	if (parseInt (app.version) > 5 && app.documents.length > 0) {
			renumber_notes ();
	}

}());