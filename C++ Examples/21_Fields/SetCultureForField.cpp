#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/SetCultureForField.docx";

	//Create a document
	intrusive_ptr<Document> document = new Document();

	//Create a section
	intrusive_ptr<Section> section = document->AddSection();

	//Add paragraph
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

	//Add textRnage for paragraph
	paragraph->AppendText(L"Add Date Field: ");

	//Add date field1
	intrusive_ptr<Field> field1 = Object::Dynamic_cast<Field>(paragraph->AppendField(L"Date1", FieldType::FieldDate));
	wstring codeString = L"DATE  \\@";
	codeString += L"\"yyyy\\MM\\dd\"";
	field1->SetCode(codeString.c_str());

	//Add new paragraph
	intrusive_ptr<Paragraph> newParagraph = section->AddParagraph();

	//Add textRnage for paragraph
	newParagraph->AppendText(L"Add Date Field with setting French Culture: ");

	//Add date field2
	intrusive_ptr<Field> field2 = newParagraph->AppendField(L"\"\\@\"dd MMMM yyyy", FieldType::FieldDate);

	//Setting Field with setting French Culture
	field2->GetCharacterFormat()->SetLocaleIdASCII(1036);

	//Update fields
	document->SetIsUpdateFields(true);

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
