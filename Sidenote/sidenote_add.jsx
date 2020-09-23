if (parseInt (app.version) > 5 && app.documents.length > 0 &&
	app.selection.length > 0 && app.selection[0].hasOwnProperty ('baseline'))
		add_sidenote (app.documents[0], 'sidenote');


function add_sidenote (doc, stylename)
	{
	var ip = app.selection[0].insertionPoints[0];
	var mode = get_mode ();
	try
		{
		var os = doc.objectStyles.item (stylename);
		var new_note = app.selection[0].textFrames.add ({appliedObjectStyle: os, label: stylename});
//~ 		var width = current_width (doc, stylename);
//~ 		if (parseInt (app.version) < 8) new_note.geometricBounds = [0, 0, "24pt", width];
		switch (mode.note_contents)
			{
			case "text_selection": {var n = app.selection[0].move (LocationOptions.atBeginning, new_note.insertionPoints[0]);
					n.applyParagraphStyle (os.appliedParagraphStyle, false)}; break;
			case "clipboard": {new_note.insertionPoints[0].select();
					app.pasteWithoutFormatting ()}; break;
			case "empty": new_note.insertionPoints[-1].select();
			}
		if (parseInt (app.version) < 8) new_note.fit (FitOptions.frameToContent);
		if (os.appliedParagraphStyle.bulletsAndNumberingListType == ListType.numberedList)
			{
			var endnote_link = doc.paragraphDestinations.add (new_note.insertionPoints[0]);
			var cr_format = doc.crossReferenceFormats.item (stylename);
			var reference = doc.crossReferenceSources.add (ip, cr_format);
			doc.hyperlinks.add (reference, endnote_link, {visible: false});
			doc.crossReferenceSources.everyItem().update();
			}
		}
	catch (e) {errorM (e.message)};
	}


function current_width (doc, stylename)
	{
	var temp = doc.objectStyles.item (stylename).label.match (/^(.+)\s(..)$/);
	try {return UnitValue (temp[1], temp[2]).as (unit_abbreviation (doc))}
		catch (e) {errorM (e.message)}
	}


function unit_abbreviation (doc)
	{
	switch (doc.viewPreferences.horizontalMeasurementUnits)
		{
		case 2051106676: errorM ("This script can't work with agataes.");
		case 2053336435: return 'cm';
		case 2053335395: return 'ci';
		case 2053729891: return 'in';
		case 2053729892: return 'in';
		case 2053991795: return 'mm';
		case 2054187363: return 'pc';
		case 2054188905: return 'pt';
		}
	}


function get_mode ()
	{
	var w = new Window ('dialog', 'Add a note', undefined, {closeButton: false});
	w.alignChildren = ['left', 'top'];
		var main = w.add ('panel');
		main.alignChildren = ['left', 'top'];
		main.orientation = 'column';
			main.add ('statictext', undefined, " 1. Add an &empty noteframe");
			main.add ('statictext', undefined, " 2. Create sidenote from &selected text");
			main.add ('statictext', undefined, " 3. Create sidenote from &clipboard contents");
		var entry = w.add ('group');
			entry.add ('staticText', undefined, "Mode: ");
			inp = entry.add ('edittext', undefined, preset (app.selection[0]));
			inp.active = true;
		var buttons = w.add ('group');
		buttons.alignment = 'right';
			buttons.orientation = 'row';
			buttons.alignChildren = ['right', 'bottom'];
			buttons.add ('button', undefined, 'OK');
			buttons.add ('button', undefined, 'Cancel', {name: 'cancel'});

	if (w.show () == 2)
		exit ();
	switch (inp.text.toUpperCase())
		{
		case "1": case "E": return {note_contents: "empty"};
		case "2": case "S": return {note_contents: "text_selection"};
		case "3": case "C": return {note_contents: "clipboard"};
		}
	}

function preset (sel)
	{
	if (sel.constructor.name == "InsertionPoint")
		return "1";
	return "2";
	}


function errorM (m)
	{
	alert (m, "Error", true)
	exit ()
	}
