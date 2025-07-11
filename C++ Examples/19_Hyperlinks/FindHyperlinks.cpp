#include "pch.h"
#include <fstream>
#include <locale>
#include <codecvt>

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/FindHyperlinks.txt";
	std::wstring inputFile = DATAPATH"/Hyperlinks.docx";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Create a hyperlink list
	std::vector<intrusive_ptr<Field>> hyperlinks;
	std::wstring hyperlinksText = L"";
	//Iterate through the items in the sections to find all hyperlinks
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
		for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++)
		{
			intrusive_ptr<DocumentObject> docObj = section->GetBody()->GetChildObjects()->GetItem(j);
			if (docObj->GetDocumentObjectType() == DocumentObjectType::Paragraph)
			{
				intrusive_ptr<Paragraph> para = Object::Dynamic_cast<Paragraph>(docObj);
				for (int k = 0; k < para->GetChildObjects()->GetCount(); k++)
				{
					intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(k);
					if (obj->GetDocumentObjectType() == DocumentObjectType::Field)
					{
						intrusive_ptr<Field> field = Object::Dynamic_cast<Field>(obj);
						if (field->GetType() == FieldType::FieldHyperlink)
						{
							hyperlinks.push_back(field);
							//Get the hyperlink text
							std::wstring text = field->GetFieldText();
							hyperlinksText.append(text.append(L"\n"));
						}
					}
				}
			}
		}
	}


	//Save the text of all hyperlinks to TXT File 
	std::wofstream write(outputFile);
	auto LocUtf8 = locale(locale(""), new std::codecvt_utf8<wchar_t>);
	write.imbue(LocUtf8);
	write << hyperlinksText;
	write.close();

	doc->Close();
}