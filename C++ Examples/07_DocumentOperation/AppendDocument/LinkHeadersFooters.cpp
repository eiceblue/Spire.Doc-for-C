#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile_1 = input_path + L"Template_N1.docx";
	wstring inputFile_2 = input_path + L"Template_N2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"LinkHeadersFooters.docx";

	//Load the source file
	intrusive_ptr<Document> srcDoc = new Document();
	srcDoc->LoadFromFile(inputFile_1.c_str());

	//Load the destination file
	intrusive_ptr<Document> dstDoc = new Document();
	dstDoc->LoadFromFile(inputFile_2.c_str());

	//Link the headers and footers in the source file
	srcDoc->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetHeader()->SetLinkToPrevious(true);
	srcDoc->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetFooter()->SetLinkToPrevious(true);

	//Clone the sections of source to destination
	int sectionCount = srcDoc->GetSections()->GetCount();
	for (int i = 0; i < sectionCount; i++)
	{
		intrusive_ptr<Section> section = srcDoc->GetSections()->GetItemInSectionCollection(i);
		dstDoc->GetSections()->Add(section->CloneSection());
	}

	//Save the document
	dstDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	srcDoc->Close();
	dstDoc->Close();
}