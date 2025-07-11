#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path+L"/example.zip";
	wstring inputFile_I = input_path+L"/example.png";
	wstring outputFile = output_path + L"InsertOLEAsIconViaStream.docx";

	//Create word document
	intrusive_ptr<Document> doc = new Document();

	//add a section
	intrusive_ptr<Section> sec = doc->AddSection();

	//add a paragraph
	intrusive_ptr<Paragraph> par = sec->AddParagraph();

	//ole stream
	intrusive_ptr<Stream> stream = new Stream(inputFile.c_str());

	//load the image
	intrusive_ptr<DocPicture> picture = new DocPicture(doc);
	picture->LoadImageSpire(inputFile_I.c_str());

	//insert the OLE from stream
	intrusive_ptr<DocOleObject> obj = par->AppendOleObject(stream, picture, L"zip");

	//display as icon
	obj->SetDisplayAsIcon(true);
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}