#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/CreateCrossReference.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Add a new section.
	intrusive_ptr<Section> section = document->AddSection();

	//Create a bookmark.
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
	paragraph->AppendBookmarkStart(L"MyBookmark");
	paragraph->AppendText(L"Text inside a bookmark");
	paragraph->AppendBookmarkEnd(L"MyBookmark");

	//Insert line breaks.
	for (int i = 0; i < 4; i++)
	{
		paragraph->AppendBreak(BreakType::LineBreak);
	}

	//Create a cross-reference field, and link it to bookmark.                    
	intrusive_ptr<Field> field = new Field(document);
	field->SetType(FieldType::FieldRef);
	field->SetCode(L"REF MyBookmark \\p \\h");

	//Insert field to paragraph.
	paragraph = section->AddParagraph();
	paragraph->AppendText(L"For more information, see ");
	paragraph->GetChildObjects()->Add(field);

	//Insert FieldSeparator object.
	intrusive_ptr<FieldMark> fieldSeparator = new FieldMark(document, FieldMarkType::FieldSeparator);
	paragraph->GetChildObjects()->Add(fieldSeparator);

	//Set display text of the field.
	intrusive_ptr<TextRange> tr = new TextRange(document);
	tr->SetText(L"above");
	paragraph->GetChildObjects()->Add(tr);

	//Insert FieldEnd object to mark the end of the field.
	intrusive_ptr<FieldMark> fieldEnd = new FieldMark(document, FieldMarkType::FieldEnd);
	paragraph->GetChildObjects()->Add(fieldEnd);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}