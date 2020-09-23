if (parseInt (app.version) > 5 && app.documents.length > 0)
	reset_sidenotes (app.documents[0]);


function reset_sidenotes (doc)
{
	//var width = doc.objectStyles.item ("sidenote").label;
	var notes = find_notes (doc);
	for (var i = 0; i < notes.length; i++)
	{
		notes[i].applyObjectStyle (doc.objectStyles.item ("sidenote"), true, false);
        // "24pt" is just an arbitrary value for the frames' height. Height is fixed by fit ()
		notes[i].geometricBounds = [0, 0, '24pt', '10pt'];
		// In CS6+ the frames use InDesign's native Autoheight
		if (parseInt (app.version) < 8)
			notes[i].fit (FitOptions.frameToContent);
	}
}


function find_notes (doc)
{
	app.findObjectPreferences = null;
	app.findObjectPreferences.appliedObjectStyles = doc.objectStyles.item ("sidenote");
	var f = app.documents[0].findObject();
	return f;
}