#include "pch.h"

using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_N3.docx";
	wstring outputFile = output_path + L"RemoveSectionContent.docx";
	//Load the Word document from disk
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Loop through all sections
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
		//Remove header content
		section->GetHeadersFooters()->GetHeader()->GetChildObjects()->Clear();
		//Remove GetBody() content
		section->GetBody()->GetChildObjects()->Clear();
		//Remove footer content
		section->GetHeadersFooters()->GetFooter()->GetChildObjects()->Clear();
	}

	//Save the Word document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}
