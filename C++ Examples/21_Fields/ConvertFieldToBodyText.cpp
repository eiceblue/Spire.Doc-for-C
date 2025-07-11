#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/ConvertFieldToBodyText.docx";
	std::wstring inputFile = DATAPATH"/TextInputField.docx";

	//Create the source document
	intrusive_ptr<Document> sourceDocument = new Document();

	//Load the source document from disk.
	sourceDocument->LoadFromFile(inputFile.c_str());

	//Traverse FormFields
	int formFieldsCount = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetFormFields()->GetCount();
	for (int i = 0; i < formFieldsCount; i++)
	{
		intrusive_ptr<FormField> field = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetFormFields()->GetItem(i);
		//Find FieldFormTextInput type field
		if (field->GetType() == FieldType::FieldFormTextInput)
		{
			//Get the paragraph
			intrusive_ptr<Paragraph> paragraph = field->GetOwnerParagraph();

			//Define variables
			int startIndex = 0;
			int endIndex = 0;

			//Create a new TextRange
			intrusive_ptr<TextRange> textRange = new TextRange(sourceDocument);

			//Set text for textRange
			textRange->SetText(paragraph->GetText());

			//Traverse DocumentObjectS of field paragraph
			int pChildObjectsCount = paragraph->GetChildObjects()->GetCount();
			for (int j = 0; j < pChildObjectsCount; j++)
			{
				intrusive_ptr<DocumentObject> obj = paragraph->GetChildObjects()->GetItem(j);
				//If its DocumentObjectType is BookmarkStart
				if (obj->GetDocumentObjectType() == DocumentObjectType::BookmarkStart)
				{
					//Get the index
					startIndex = paragraph->GetChildObjects()->IndexOf(obj);
				}
				//If its DocumentObjectType is BookmarkEnd
				if (obj->GetDocumentObjectType() == DocumentObjectType::BookmarkEnd)
				{
					//Get the index
					endIndex = paragraph->GetChildObjects()->IndexOf(obj);
				}
			}
			//Remove ChildObjects
			for (int k = endIndex; k > startIndex; k--)
			{
				//If it is TextFormField
				if (Object::CheckType<TextFormField>(paragraph->GetChildObjects()->GetItem(k)))
				{
					intrusive_ptr<TextFormField> textFormField = boost::dynamic_pointer_cast<TextFormField>(paragraph->GetChildObjects()->GetItem(k));

					//Remove the field object
					paragraph->GetChildObjects()->Remove(textFormField);
				}
				else
				{
					paragraph->GetChildObjects()->RemoveAt(k);
				}
			}
			//Insert the new TextRange
			paragraph->GetChildObjects()->Insert(startIndex, textRange);

			break;
		}

	}
	//Save the document.
	sourceDocument->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	sourceDocument->Close();
}