#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertPageBreakSecondApproach.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Insert page break.
	document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(3)->AppendBreak(BreakType::PageBreak);

	//Save the file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}
