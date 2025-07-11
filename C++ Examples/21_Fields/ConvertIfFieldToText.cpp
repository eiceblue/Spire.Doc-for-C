#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/ConvertIfFieldToText.docx";
	std::wstring inputFile = DATAPATH"/IfFieldSample.docx";

	//Open a Word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get all fields in document
	intrusive_ptr<FieldCollection> fields = document->GetFields();

	for (int i = 0; i < fields->GetCount(); i++)
	{
		intrusive_ptr<Field> field = fields->GetItem(i);
		if (field->GetType() == FieldType::FieldIf)
		{
			intrusive_ptr<TextRange> original = Object::Dynamic_cast<TextRange>(field);
			//Get field text
			std::wstring text = field->GetFieldText();
			//Create a new textRange and set its format
			intrusive_ptr<TextRange> textRange = new TextRange(document);
			textRange->SetText(text.c_str());
			textRange->GetCharacterFormat()->SetFontName(original->GetCharacterFormat()->GetFontName());
			textRange->GetCharacterFormat()->SetFontSize(original->GetCharacterFormat()->GetFontSize());

			intrusive_ptr<Paragraph> par = field->GetOwnerParagraph();
			//Get the index of the if field
			int index = par->GetChildObjects()->IndexOf(field);
			//Remove if field via index
			par->GetChildObjects()->RemoveAt(index);
			//Insert field text at the position of if field
			par->GetChildObjects()->Insert(index, textRange);
		}
	}
	//Save doc file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}