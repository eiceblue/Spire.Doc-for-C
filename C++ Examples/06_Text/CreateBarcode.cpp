#include "pch.h"


using namespace Spire::Doc;

int main()
{
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreateBarcode.docx";

	//Create a document
	intrusive_ptr<Document> doc =  new Document();

	//Add a paragraph
	intrusive_ptr<Paragraph> p = doc->AddSection()->AddParagraph();

	//Add barcode and set its format
	intrusive_ptr<TextRange> txtRang = p->AppendText(L"H63TWX11072");
	//Set barcode font name, note you need to install the barcode font on your system at first
	txtRang->GetCharacterFormat()->SetFontName(L"C39HrP60DlTt");
	txtRang->GetCharacterFormat()->SetFontSize(80);
	txtRang->GetCharacterFormat()->SetTextColor(Color::GetSeaGreen());

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}