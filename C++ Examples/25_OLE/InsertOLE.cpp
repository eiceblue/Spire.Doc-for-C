#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertOLE.docx";

	//create a document
	intrusive_ptr<Document> doc = new Document();

	//add a section
	intrusive_ptr<Section> sec = doc->AddSection();

	//add a paragraph
	intrusive_ptr<Paragraph> par = sec->AddParagraph();

	//load the image
	intrusive_ptr<DocPicture> picture = new DocPicture(doc);
	wstring imagePath = input_path + L"Excel.png";
	picture->LoadImageSpire(imagePath.c_str());

	//insert the OLE
	wstring filePath = input_path + L"example.xlsx";
	intrusive_ptr<DocOleObject> obj = par->AppendOleObject(filePath.c_str(), picture, OleObjectType::ExcelWorksheet);
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}
