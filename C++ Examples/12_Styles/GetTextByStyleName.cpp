#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_N5.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetTextByStyleName.txt";

	//Load document from disk
	intrusive_ptr<Document>  doc = new Document();
	doc->LoadFromFile(inputFile.c_str());


	//Create string builder
	wstring* builder = new wstring();
	//Loop through sections
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
		//Loop through paragraphs
		for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
		{
			intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(j);
			//Find the paragraph whose style name is "Heading1"
			wstring style_name = para->GetStyleName();
			if (style_name.compare(L"Heading1") == 0)
			{
				//Write the text of paragraph
				builder->append(para->GetText());
				builder->append(L"\n");
			}
		}
	}

	//Write the contents in a TXT file
	std::wofstream foo(outputFile);
	foo << builder->c_str();
	foo.close();
	doc->Close();
	delete builder;
}
