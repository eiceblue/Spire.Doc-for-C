#include "pch.h"


using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring outputFile = output_path + L"SetSpacing.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());
	//Add the text strings to the paragraph and set the style.
	intrusive_ptr<Paragraph> para = new Paragraph(document);
	intrusive_ptr<TextRange> textRange1 = para->AppendText(L"This is an inserted paragraph.");
	textRange1->GetCharacterFormat()->SetTextColor(Color::GetBlue());
	textRange1->GetCharacterFormat()->SetFontSize(15);

	//set the spacing before and after.
	para->GetFormat()->SetBeforeAutoSpacing(false);
	para->GetFormat()->SetBeforeSpacing(10);
	para->GetFormat()->SetAfterAutoSpacing(false);
	para->GetFormat()->SetAfterSpacing(10);

	//insert the added paragraph to the word document.
	document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->Insert(1, para);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}