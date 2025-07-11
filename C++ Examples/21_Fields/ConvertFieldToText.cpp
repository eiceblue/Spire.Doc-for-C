#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring inputFile = DATAPATH"/Fields.docx";
	std::wstring outputFile = OUTPUTPATH"/ConvertFieldToText.docx";

	//Load word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get all fields in document
	intrusive_ptr<FieldCollection> fields = document->GetFields();
	int count = fields->GetCount();

	for (int i = 0; i < count; i++)
	{
		intrusive_ptr<Field> field = fields->GetItem(0);
		std::wstring s = field->GetFieldText();
		int index = field->GetOwnerParagraph()->GetChildObjects()->IndexOf(field);
		intrusive_ptr<TextRange> textRange = new TextRange(document);
		textRange->SetText(s.c_str());
		textRange->GetCharacterFormat()->SetFontSize(24.0f);

		field->GetOwnerParagraph()->GetChildObjects()->Insert(index, textRange);
		field->GetOwnerParagraph()->GetChildObjects()->Remove(field);
	}

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}