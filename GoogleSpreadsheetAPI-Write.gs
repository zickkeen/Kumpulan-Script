function doPost(e) {
	var ss = SpreadsheetApp.openById("1pKKGTFVN4L-onBzmJqk4gmVI4rTFEW17IZr4El1OgVs");
    var sheet = ss.getSheetByName("link");
	var action = e.parameter.action;

	switch(action){
		case "insert":
			return insert_data(e,sheet);
			break;
	}
  }

function insert_data(request, sheet){
	var waktu = request.parameter.waktu;
	var judul = request.parameter.judul;
	var alamat = request.parameter.alamat;
	var tipe = request.parameter.tipe;

	sheet.appendRow([waktu,judul,alamat,tipe]);
	var hasil = "Data berhasil diinput";

	hasil = JSON.stringify(
		{
          "hasil" : hasil
		}
	);
	
	return ContentService.createTextOutput(hasil).setMimeType(ContentService.MimeType.JSON);

}
