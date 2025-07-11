#include "pch.h"
#include <fstream>
#include <locale>
#include <codecvt>


using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ExtractParagraphBasedOnStyle.docx";
	wstring outputFile = output_path + L"ExtractParagraphBasedOnStyle.txt";

	//Create a new document
	intrusive_ptr<Document> document = new Document();
	wstring styleName1 = L"Heading1";
	wstring style1Text;
	//Load file from disk
	document->LoadFromFile(inputFile.c_str());
	style1Text.append(L"The following is the content of the paragraph with the style name " + styleName1 + L": \n");
	//Extrct paragraph based on style
	for (int i = 0; i < document->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
		for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
		{
			intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetItemInParagraphCollection(j);
			if (paragraph->GetStyleName() != nullptr && paragraph->GetStyleName() == styleName1)
			{
				style1Text.append(paragraph->GetText());
			}
		}
	}

	std::wofstream write(outputFile);
	auto LocUtf8 = locale(locale(""), new std::codecvt_utf8<wchar_t>);
	write.imbue(LocUtf8);
	write << style1Text;
	write.close();

	//File::WriteAllText(outputFile.c_str(), style1Text->toString());
	document->Close();
}

