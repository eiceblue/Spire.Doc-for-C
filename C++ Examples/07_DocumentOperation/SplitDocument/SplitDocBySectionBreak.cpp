#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_4.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SplitDocBySectionBreak/";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Define another new word document object.
	intrusive_ptr<Document> newWord;

	//Split a Word document into multiple documents by section break.
	for (int i = 0; i < document->GetSections()->GetCount(); i++)
	{
		std::wstring result = outputFile.c_str();
		result += L"SplitDocBySectionBreak_" + to_wstring(i) + L".docx";
		newWord = new Document();
		newWord->GetSections()->Add(document->GetSections()->GetItemInSectionCollection(i)->CloneSection());

		//Save to file.
		newWord->SaveToFile(result.c_str());
		newWord->Close();
	}

}