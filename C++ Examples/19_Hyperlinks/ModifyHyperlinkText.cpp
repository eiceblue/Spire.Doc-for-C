#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/ModifyHyperlinkText.docx";
	std::wstring inputFile = DATAPATH"/Hyperlinks.docx";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	std::vector<intrusive_ptr<Field>> hyperlinks;
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
						}
					}
				}
			}
		}
	}

	//Reset the property of hyperlinks->GetItem(0)->FieldText by using the index of the hyperlink
	hyperlinks[0]->SetFieldText(L"Spire.Doc component");

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}