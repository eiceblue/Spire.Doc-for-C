#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_4.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddPageNumbersInSections.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Repeat step2 and Step3 for the rest sections, so change the code with for loop.
	for (int i = 0; i < 3; i++)
	{
		intrusive_ptr<HeaderFooter> footer = document->GetSections()->GetItemInSectionCollection(i)->GetHeadersFooters()->GetFooter();
		intrusive_ptr<Paragraph> footerParagraph = footer->AddParagraph();
		footerParagraph->AppendField(L"page number", FieldType::FieldPage);
		footerParagraph->AppendText(L" of ");
		footerParagraph->AppendField(L"number of pages", FieldType::FieldSectionPages);
		footerParagraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

		if (i == 2)
		{
			break;
		}
		else
		{
			document->GetSections()->GetItemInSectionCollection(i + 1)->GetPageSetup()->SetRestartPageNumbering(true);
			document->GetSections()->GetItemInSectionCollection(i + 1)->GetPageSetup()->SetPageStartingNumber(1);
		}
	}
	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}
