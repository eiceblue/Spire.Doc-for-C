#include "pch.h"
using namespace Spire::Doc;

int main() {
	std::wstring inputFile = DATAPATH"/ComboBox.docx";
	std::wstring outputFile = OUTPUTPATH"/ComboBoxItem.docx";
	
	// Create a new document and load from file
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the combo box from the file
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
		for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++)
		{
			intrusive_ptr<DocumentObject> bodyObj = section->GetBody()->GetChildObjects()->GetItem(j);
			if (bodyObj->GetDocumentObjectType() == DocumentObjectType::StructureDocumentTag)
			{
				//If SDTType is ComboBox
				if ((Object::Dynamic_cast<StructureDocumentTag>(bodyObj))->GetSDTProperties()->GetSDTType() == SdtType::ComboBox)
				{
					intrusive_ptr<StructureDocumentTag> sdt = Object::Dynamic_cast<StructureDocumentTag>(bodyObj);
					intrusive_ptr<SdtControlProperties> scp = sdt->GetSDTProperties()->GetControlProperties();
				
					intrusive_ptr<SdtComboBox> combo = Object::Convert<SdtComboBox>(scp);
					//Remove the second list item
					combo->GetListItems()->RemoveAt(1);
					//Add a new item
					intrusive_ptr<SdtListItem> item = new SdtListItem(L"D", L"D");
					combo->GetListItems()->Add(item);

					//If the value of list items is "D"
					for (int i = 0; i < combo->GetListItems()->GetCount(); i++)
					{
						intrusive_ptr<SdtListItem> sdtItem = combo->GetListItems()->GetItem(i);
						if (wcscmp(sdtItem->GetValue(), L"D") == 0)
						{
							//Select the item
							combo->GetListItems()->SetSelectedValue(sdtItem);
						}
					}
				}
			}
		}
	}

	//Save the document and launch it
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}
