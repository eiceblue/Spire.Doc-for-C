#include "pch.h"


using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring outputFile = output_path + L"ManagePagination.docx";

	//Create Word document.
	intrusive_ptr<Document> document =new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Get the first section and the paragraph we want to manage the pagination.
	intrusive_ptr<Section> sec = document->GetSections()->GetItemInSectionCollection(0);
	intrusive_ptr<Paragraph> para = sec->GetParagraphs()->GetItemInParagraphCollection(4);

	//Set the pagination format as Format.PageBreakBefore for the checked paragraph.
	para->GetFormat()->SetPageBreakBefore(true);

	//Save the file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}