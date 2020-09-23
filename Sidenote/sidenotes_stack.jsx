if (parseInt (app.version) > 5 && app.documents.length > 0)
{
	try {stack_notes ()}
		catch (e) {alert (e.message + "\r(line " + e.line + ")")};
}

// =============================================================================

function stack_notes ()
{
	var stack = get_position ();
	var units = app.scriptPreferences.measurementUnit;
	app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
	app.documents[0].zeroPoint = [0,0];
	switch (stack.target) {
		case "document": stack_document (stack); break;
		case "textframe": stack_frame (app.selection[0], stack); break
	}
	app.scriptPreferences.measurementUnit = units;
}


function stack_document (stack)
{
	var win = create_message (40, "Stack");
	win.show();
	win.message.text = 'Stacking...';
	var frame_IDs = find_noted_frames (stack.stylename);
	var textframe;
	for (var i = 0; i < frame_IDs.length; i++)
	{
		win.message.text = 'Stacking...' + String(i);
		textframe = app.documents[0].textFrames.itemByID (frame_IDs[i]);
		stack_frame (textframe, stack)
	}
	try {win.message.parent.close()} catch(_){}
}


function stack_frame (tf, stack)
{
	stack.notes = find_notes (tf, stack.stylename);
	var block_height = getStackedHeight (stack);
	stack.main_frame = stack.notes[0].parent.parentTextFrames[0];
	stack.main_bounds = stack.main_frame.geometricBounds;
	switch (stack.position)
	{
		case "top": stack.start = stack.main_bounds[0]; break;
		case "centre": stack.start = stack.main_bounds[0] + (((stack.main_bounds[2]-stack.main_bounds[0])/2) - (block_height/2)); break;
		case "bottom": stack.start = stack.main_bounds[2] - block_height;
	}
	place_notes (stack);
}


function place_notes (stack) {
	var x = stack.notes[0].geometricBounds[1];
	yPos (stack.notes[0]);
	stack.notes[0].move ([x, stack.start]);
	
	if (stack.baseline){
		adjust_first_note (stack);
	}

	for (var i = 1; i < stack.notes.length; i++) {
		yPos (stack.notes[i]);
		stack.notes[i].move ([x, stack.notes[i-1].geometricBounds[2] + stack.space]);
	}
}



function adjust_first_note (stack) {
	if (stack.position == 'top') {
		var diff = stack.main_frame.lines[0].baseline - stack.notes[0].lines[0].baseline;
		stack.notes[0].move (undefined, [0, diff]);
	} else {
		//var diff = stack.main_frame.geometricBounds[2] - stack.main_frame.lines[-1].baseline;
		var diff = stack.main_frame.geometricBounds[2] - getLastBaseline (stack.main_frame);
		stack.notes[0].move (undefined, [0, -diff]);
	}
}


function find_notes (tf, stylename){
	var notes = [];
	var anchors = tf.allPageItems;
	for (var i = 0; i < anchors.length; i++){
		if (anchors[i].appliedObjectStyle.name == stylename){
			notes.push (anchors[i]);
		}
	}
	return notes;
}


function getStackedHeight (stack){
	var h = height (stack.notes[0]);
	for (var i = stack.notes.length-1; i > 0; i--){
		h += height (stack.notes[i]) + stack.space;
	}
	return h;
}


function height (frame) {
	return frame.geometricBounds[2]-frame.geometricBounds[0];
}


function yPos (frame) {
	frame.anchoredObjectSettings.anchorYoffset = 0;
	
	if (frame.anchoredObjectSettings.verticalReferencePoint != VerticallyRelativeTo.TEXT_FRAME) {
		frame.anchoredObjectSettings.verticalReferencePoint = VerticallyRelativeTo.TEXT_FRAME;
	}
}


function getLastBaseline (frame) {
	
		function deepestCellInRow (row) {
			var baselines = row.cells.everyItem().lines[-1].baseline;
			return Math.max.apply (null, baselines);
		}

	if (frame.contents.length === 0 && frame.paragraphs.length > 0 && frame.paragraphs[0].lines.length > 1) {
		return deepestCellInRow (frame.paragraphs[0].tables[-1].rows[-1]);
	}
	
	if (frame.paragraphs[-1].tables.length === 0) {
		return frame.lines[-1].baseline;
	}

	if (frame.paragraphs[-1].lines.length === 1) { // Table doesn't break
		return frame.paragraphs[-1].tables[0].rows[-1].lines[-1].baseline;
	}

	var n = 0;
	var cells = frame.paragraphs[-1].tables[0].columns[0].cells.everyItem().getElements();
	var stop = cells.length;
	while (cells[n].insertionPoints[0].parentTextFrames[0] == cells[n+1].insertionPoints[0].parentTextFrames[0]
				&& n < stop) {
		n++;
	}
	
	return deepestCellInRow (cells[n].parentRow);
}


function find_noted_frames (stylename) {
	var parent_id;
	app.findObjectPreferences = null;
	app.findObjectPreferences.appliedObjectStyles = app.documents[0].objectStyles.item (stylename);
	var f = app.documents[0].findObject();
	var known = [];
	var list = [];
	for (var i = 0; i < f.length; i++){
		try {
			parent_id = f[i].parent.parentTextFrames[0].id;
			if (!known[parent_id]){
				list.push (parent_id);
				known[parent_id] = true;
			}
		} catch (_){}
	}
	return list;
}


function get_position (){
	var prefs = read_settings();
	var w = new Window ('dialog', 'Stack sidenotes', undefined, {closeButton: false});
		var main = w.add ('panel'); main.alignChildren = 'fill';
		
		var g1 = main.add ('panel {orientation: "column", alignChildren: "left"}');
			var vpos = [];
			vpos[0] = g1.add ('radiobutton {text: "Top"}');
			vpos[1] = g1.add ('radiobutton {text: "Centre"}');
			vpos[2] = g1.add ('radiobutton {text: "Bottom"}');
			vpos[prefs.position].value = true;
			
		var g2 = main.add ('panel {orientation: "column", alignChildren: "left"}');
			var scope = [];
			scope[0] = g2.add ('radiobutton {text: "Document"}');
			scope[1] = g2.add ('radiobutton {text: "Selected text frame only"}');
			
		var sp = main.add ('group');
			sp.add ('statictext {text: "Space between notes: "}');
			space = sp.add ('edittext {characters: 10}'); //, undefined, convert_units ("6pt", doc_units ()));
			space.characters = 10;
			
		var baseline = main.add ('checkbox {text: "Align baselines"}');
		baseline.helpTip = 'Align the first baseline of the first note with the first baseline of the text (or the last baseline of the last note with the last baseline of the text)'
			
		var os = main.add ("group {orientation: 'row'}");
			os.add ('statictext {text: "Style: "}');
			var ostyles = os.add ("dropdownlist", undefined, ostyle_names (app.documents[0]));
				ostyles.minimumSize.width = ostyles.maximumSize.width = 180;
				
		var buttons = w.add ('group {alignment: "right", orientation: "row", alignChildren: ["right", "bottom"]}');
			buttons.add ('button', undefined, 'OK');
			buttons.add ('button', undefined, 'Cancel', {name: 'cancel'});

		var default_style = ostyles.find ("sidenote");
		
		vpos[0].onClick = function(){baseline.enabled = true;}
		vpos[1].onClick = function(){baseline.enabled = baseline.value = false;}
		vpos[2].onClick = function(){baseline.enabled = true;}
		
		if (default_style == null){
			ostyles.selection = 0;
			space.text = convert_units ("6pt", doc_units ());
		}else{
			ostyles.selection = default_style;
			var previous_space = app.documents[0].objectStyles.item ('sidenote').extractLabel ('sidenote_vspace');
			if (previous_space == ""){
				space.text = convert_units ("6pt", doc_units ());
			}else{
				space.text = convert_units (previous_space, doc_units());
			}
		}

	if (app.selection.length > 0 && app.selection[0] instanceof TextFrame && app.selection[0].textFrames.length > 0) {
		scope[1].value = true;
	} else {
		scope[0].value = true;
	}
	
	baseline.value = prefs.baseline;
	
	space.onChange = function () {space.text = convert_units (space.text, doc_units ())}
	
	if (w.show() == 2) exit ();
	
	app.documents[0].objectStyles.item (ostyles.selection.text).insertLabel ('sidenote_vspace', space.text);
	
	var obj = {
		position: array_index (vpos),
		baseline: baseline.value
	}

	write_settings (obj);
	
	return {
		position: ["top", "centre", "bottom"][array_index (vpos)],
		target: ["document", "textframe"][array_index (scope)],
		space: Number (convert_units (space.text, "pt").replace (/\spt$/, "")),
		stylename: ostyles.selection.text,
		baseline: baseline.value
		}
	}


function read_settings () {
	var obj = {position: 2, baseline: false} // Some defaults
	var f = File (script_dir()+'/sidenotes_stack.txt');
	if (f.exists){
		try{
			f.open ('r');
			var obj = eval (f.read());
		} catch (_){}
	}
	return obj;
}


function write_settings (obj) {
	var f = File (script_dir()+'/sidenotes_stack.txt');
	f.open ('w');
	f.write (obj.toSource ());
	f.close ()
}


function script_dir(){
	try {return File (app.activeScript).path}
	catch(e) {return File (e.fileName).path}
}


function ostyle_names (doc){
	var array = [];
	var os = doc.objectStyles.everyItem().name;
	for (var i = 0; i < os.length; i++){
		if (os[i].charAt (0) != "["){
			array.push (os[i]);
		}
	}
	return array;
}


function array_index (array){
	for (var i = 0; i < array.length; i++){
		if (array[i].value == true){
			return i;
		}
	}
}

//--------------------------------------------------------------------
function convert_units (n, to){
	var unitConversions = {
		'pt': 1.0000000000,
		'p': 12.0000000000,
		'mm': 2.8346456692,
		'in': 72.00000000,
		'ag': 5.1428571428,
		'cm': 28.3464566929,
		'c': 12.7878751998,
		'tr': 3.0112500000 // traditional point -- but we don't do anything with it yet
	}
	var obj = fix_measurement (n);
	var temp = (obj.amount * unitConversions[obj.unit]) / unitConversions[to];
	return output_format (temp, to)
}

// Add the target unit to the amount, either suffixed pt, ag, mm, cm, in, or infixed p or c

function output_format (amount, target) {
	amount = amount.toFixed(3).replace(/\.?0+$/, '');
	if (target.length == 2) { // two-character unit: pt, mm, etc.
		return String (amount) + ' ' + target;
	} else {// 'p' or 'c'
		var decimal = (Number (amount) - parseInt (amount)) * 12;
		return parseInt (amount) + target + decimal;
		}
}


function fix_measurement (n) {
	n = n.replace(/ /g,'');
	n = n.replace (/(\d+)([pc])([.\d]+)$/, function () {return Number (arguments[1]) + Number (arguments[3]/12) + arguments[2]});
	var temp = n.split (/(ag|cm|mm|c|pt|p|in)$/);
	return {amount: Number (temp[0]), unit: temp.length === 1 ? doc_units() : temp[1]};
}


function doc_units () {
	switch (app.documents[0].viewPreferences.horizontalMeasurementUnits){
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


function errorM (m) {alert (m); exit ()}


function create_message (le, title)
{
	var w = new Window ("palette", title);
	w.alignChildren = ["left", "top"];
	w.message = w.add ("statictext", undefined, "");
	w.message.characters = le;
	return w;
}