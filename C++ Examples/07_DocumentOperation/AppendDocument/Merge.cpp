#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile_1 = input_path + L"Summary_of_Science.doc";
	wstring inputFile_2 = input_path + L"Bookmark.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Merge.docx";

	//Create word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile_1.c_str(), FileFormat::Doc);

	intrusive_ptr<Document> documentMerge = new Document();
	documentMerge->LoadFromFile(inputFile_2.c_str(), FileFormat::Docx);


	int sectionCount = documentMerge->GetSections()->GetCount();
	for (int i = 0; i < sectionCount; i++)
	{
		intrusive_ptr<Section> section = documentMerge->GetSections()->GetItemInSectionCollection(i);
		document->GetSections()->Add(section->CloneSection());
	}

	//Save as docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	documentMerge->Close();
}