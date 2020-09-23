if(app.documents.length>0)
{
	var ad = app.activeDocument;
	var tf = ad.textFrames;
	var tflg = tf.length;
	if(tflg>0)
	{
		var wcount = 0;
		var chcount = 0;
		var pcount=0;

		for(i=0; i<tflg; i++)
		{	
			var p = tf[i].paragraphs;
			for(l=0; l<p.length; l++)
			{
				pcount+=1;
				wcount += p[l].words.length;
				chcount += p[l].characters.length;
			}
		}

		alert("Açık Olan Dökümanda: "+"\r"
        	+ "- "+tflg+ " Yazı Alanı" + "\r"
			+ "- "+pcount + " Paragraf" + "\r"
			+ "- " +wcount + " Kelime" + "\r"  
			+ "- "+chcount + " Karakter (Boşluklar dahil)" + "\r" 
			+ "- "+(chcount-spaced()) + " Karakter (Boşluklar Hariç)" + "\r"
              + "BULUNMAKTADIR.");
	}
}

function spaced()
{
	app.findGrepPreferences = app.changeGrepPreferences = null;
	app.findGrepPreferences.findWhat="\s";
	return app.activeDocument.findGrep().length;
}