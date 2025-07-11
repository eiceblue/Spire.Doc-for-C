#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"MultiStylesInAParagraph.docx";

	//Create a Word document
	intrusive_ptr<Document> doc = new Document();

	//Add a section
	intrusive_ptr<Section> section = doc->AddSection();

	//Add a paragraph
	intrusive_ptr<Paragraph> para = section->AddParagraph();

	//Add a text range 1 and set its style
	intrusive_ptr<TextRange> range = para->AppendText(L"Spire.Doc for .NET ");
	range->GetCharacterFormat()->SetFontName(L"Calibri");
	range->GetCharacterFormat()->SetFontSize(16.0f);
	range->GetCharacterFormat()->SetTextColor(Color::GetBlue());
	range->GetCharacterFormat()->SetBold(true);
	range->GetCharacterFormat()->SetUnderlineStyle(UnderlineStyle::Single);

	//Add a text range 2 and set its style
	range = para->AppendText(L"is a professional Word .NET library");
	range->GetCharacterFormat()->SetFontName(L"Calibri");
	range->GetCharacterFormat()->SetFontSize(15.0f);

	//Save the Word document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}