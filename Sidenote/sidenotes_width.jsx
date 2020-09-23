//DESCRIPTION: set the width of all sidenotes
// Script assumes presence of styles/items named "sidenote"
// Peter Kahrel -- www.kahrel.plus.com

#target indesign;

if (parseInt (app.version) > 5 && app.documents.length > 0)
	sidenote_width (app.documents[0], "sidenote");


function sidenote_width (doc, stylename)
	{
	if (doc.objectStyles.item (stylename) == null)
		errorM ("Object style '" + stylename + "' not found.");
	try
		{
		var width = sidenote_coordinates (doc.objectStyles.item (stylename));
		doc.objectStyles.item (stylename).textFramePreferences.textColumnFixedWidth = width;
		if (parseInt (app.version) < 8)
			{
			var p_items = doc.allPageItems;
			var notes = find_notes (doc);
			pbar____ = progress_bar (notes.length, "Fitting sidenotes")
			for (var i = 0; i < notes.length; i++)
				{
				pbar____.value = i;
	//~ 			gb = notes[i].geometricBounds;
	//~ 			notes[i].geometricBounds = [0, 0, gb[2], width];
				notes[i].fit (FitOptions.frameToContent);
				}
	//~ 		doc.objectStyles.item (stylename).label = width;
			}
		} catch (e) {errorM (e.message)}
	try {pbar____.parent.close ();} catch (_) {}
	}


function find_notes (doc)
	{
	app.findObjectPreferences = null;
	app.findObjectPreferences.appliedObjectStyles = doc.objectStyles.item ("sidenote");
	return app.documents[0].findObject();
	}

// Interface ==========================================================================

function sidenote_coordinates (objStyle)
	{
	var w = new Window ('dialog', 'Resize sidenotes', undefined, {closeButton: false});
	var doc_unit = doc_units ();
	w.alignChildren = ['right', 'top'];
		var main = w.add ('panel');
		main.alignChildren = ['right', 'top'];
			var g1 = main.add ('group');
				g1.add ('statictext', undefined, 'Width of the notes: ');
				var width = g1.add ('edittext');
				width.characters = 10;
//~ 				width.text = convert_units (previous, doc_unit);
				width.text = convert_units (String (objStyle.textFramePreferences.textColumnFixedWidth), doc_unit);
				width.active = true;
		var buttons = w.add ('group');
			buttons.orientation = 'row';
			buttons.alignChildren = ['right', 'bottom'];
			buttons.add ('button', undefined, 'OK');
			buttons.add ('button', undefined, 'Cancel', {name: 'cancel'});

		width.onChange = function () {width.text = convert_units (width.text, doc_unit)};
		
	if (w.show () == 2)
		exit ();
	return width.text
	}



function convert_units (n, to)
	{
	var m = [];
	m["ag"] = 5.1428571428;
	m["p"] = 12.0000000000;
	m["mm"] = 2.8346456692;
	m["cm"] = 28.3464566929;
	m["in"] = 72.00000000;
	m["c"] = 12.7878751998;
	m["tr"] = 3.0112500000; // traditional point -- but we don't do anything with it yet
	m["pt"] = 1.0000000000;
	obj = fix_measurement (n);
	var temp = (obj.amount * m[obj.unit]) / m[to];
	return output_format (temp, to)
	}


// Add the target unit to the amount, either suffixed pt, ag, mm, cm, in,
// or infixed p or c

function output_format (amount, target)
	{
	amount = amount.toFixed (3).replace (/\.?0+$/g, "");
	if (target.length == 2) // two-character unit: pt, mm, etc.
		return String (amount) + " " + target;
	else // "p" or "c"
		{
		// calculate the decimal
		var decimal = (Number (amount) - parseInt (amount)) * 12;
		// return the integer part of the result + infix + formatted decimal
		return parseInt (amount) + target + decimal;
		}
	}


function fix_measurement (n)
	{
	// infixed "p" and "c" to decimal suffixes: 3p4 > 3.5 p
	n = n.replace (/(\d+)([pc])([.\d]+)$/, function () {return Number (arguments[1]) + Number (arguments[3]/12) + arguments[2]});
	// add unit if necessary
	n = n.replace (/(\d)$/, "$1" + doc_units (app.documents[0]))
	// split on unit
	var temp = n.split (/(ag|cm|mm|c|pt|p|in)$/);
	if (temp.length == 1)
		return {amount: Number (temp[0]), unit: doc_units ()};
	else
		return {amount: Number (temp[0]), unit: temp[1]};
	}


function doc_units ()
	{
	switch (app.documents[0].viewPreferences.horizontalMeasurementUnits)
		{
		case 2051106676: return 'ag';
		case 2053336435: return 'cm';
		case 2053335395: return 'c';
		case 2053729891: return 'in';
		case 2053729892: return 'in';
		case 2053991795: return 'mm';
		case 2054187363: return 'p';
		case 2054188905: return 'pt';
		}
	}


function errorM (m)
	{
	alert (m, "Error", true)
	exit ()
	}


function progress_bar (stop, title)
	{
	progressw = Window.find ('paletter', title);
	if (progressw === null)
		{
		progressw = new Window ('palette', title);
		pb____ = progressw.add ('progressbar', undefined, 0, stop);
		pb____.preferredSize = [300,20];
		}
	progressw.show()
	return pb____;
	}