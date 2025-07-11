#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Styles.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RetrieveStyle.txt";

	//Load a template document 
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Traverse all paragraphs in the document and get their style names through StyleName property
	std::wstring styleName = L"";
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
		for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
		{
			intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetItemInParagraphCollection(j);
			styleName.append(paragraph->GetStyleName());
			styleName.append(L"\n");
		}
	}

	std::wofstream foo(outputFile);
	foo << styleName;
	foo.close();
	doc->Close();
}