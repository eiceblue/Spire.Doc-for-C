#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Sample_two sections.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SectionBreakContinuous.docx";

	//Open a Word document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	int sectionCount = doc->GetSections()->GetCount();
	for (int i = 0; i < sectionCount; i++)
	{
		intrusive_ptr<Section> sec = doc->GetSections()->GetItemInSectionCollection(i);
		//Set section break as continuous
		sec->SetBreakCode(SectionBreakType::NoBreak);
	}

	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}