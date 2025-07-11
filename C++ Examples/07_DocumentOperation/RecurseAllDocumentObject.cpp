#include "pch.h"
#include <locale>
#include <codecvt>

using namespace Spire::Doc;

wstring GetDocumentObjectType(DocumentObjectType type)
{
	switch (type)
	{
	case DocumentObjectType::Any:
		return L"Any";
		break;
	case DocumentObjectType::Body:
		return L"Body";
		break;
	case DocumentObjectType::BookmarkEnd:
		return L"BookmarkEnd";
		break;
	case DocumentObjectType::BookmarkStart:
		return L"BookmarkStart";
		break;
	case DocumentObjectType::Break:
		return L"Break";
		break;
	case DocumentObjectType::BuildingBlock:
		return L"BuildingBlock";
		break;
	case DocumentObjectType::CheckBox:
		return L"CheckBox";
		break;
	case DocumentObjectType::Comment:
		return L"Comment";
		break;
	case DocumentObjectType::CommentMark:
		return L"CommentMark";
		break;
	case DocumentObjectType::ControlField:
		return L"ControlField";
		break;
	case DocumentObjectType::CustomXml:
		return L"CustomXml";
		break;
	case DocumentObjectType::Document:
		return L"Document";
		break;
	case DocumentObjectType::DropDownFormField:
		return L"DropDownFormField";
		break;
	case DocumentObjectType::EmbededField:
		return L"EmbededField";
		break;
	case DocumentObjectType::Field:
		return L"Field";
		break;
	case DocumentObjectType::FieldEnd:
		return L"FieldEnd";
		break;
	case DocumentObjectType::FieldMark:
		return L"FieldMark";
		break;
	case DocumentObjectType::FieldSeparator:
		return L"FieldSeparator";
		break;
	case DocumentObjectType::FieldStart:
		return L"FieldStart";
		break;
	case DocumentObjectType::Footnote:
		return L"Footnote";
		break;
	case DocumentObjectType::FormField:
		return L"FormField";
		break;
	case DocumentObjectType::GlossaryDocument:
		return L"GlossaryDocument";
		break;
	case DocumentObjectType::HeaderFooter:
		return L"HeaderFooter";
		break;
	case DocumentObjectType::MergeField:
		return L"MergeField";
		break;
	case DocumentObjectType::MoveFromRangeEnd:
		return L"MoveFromRangeEnd";
		break;
	case DocumentObjectType::MoveFromRangeStart:
		return L"MoveFromRangeStart";
		break;
	case DocumentObjectType::MoveToRangeEnd:
		return L"MoveToRangeEnd";
		break;
	case DocumentObjectType::MoveToRangeStart:
		return L"MoveToRangeStart";
		break;
	case DocumentObjectType::OfficeMath:
		return L"OfficeMath";
		break;
	case DocumentObjectType::OleObject:
		return L"OleObject";
		break;
	case DocumentObjectType::Paragraph:
		return L"Paragraph";
		break;
	case DocumentObjectType::PermissionEnd:
		return L"PermissionEnd";
		break;
	case DocumentObjectType::PermissionStart:
		return L"PermissionStart";
		break;
	case DocumentObjectType::Picture:
		return L"Picture";
		break;
	case DocumentObjectType::Ruby:
		return L"Ruby";
		break;
	case DocumentObjectType::SDTBlockContent:
		return L"SDTBlockContent";
		break;
	case DocumentObjectType::SDTCellContent:
		return L"SDTCellContent";
		break;
	case DocumentObjectType::SDTInlineContent:
		return L"SDTInlineContent";
		break;
	case DocumentObjectType::SDTRowContent:
		return L"SDTRowContent";
		break;
	case DocumentObjectType::Section:
		return L"Section";
		break;
	case DocumentObjectType::SeqField:
		return L"SeqField";
		break;
	case DocumentObjectType::Shape:
		return L"Shape";
		break;
	case DocumentObjectType::ShapeGroup:
		return L"ShapeGroup";
		break;
	case DocumentObjectType::ShapeLine:
		return L"ShapeLine";
		break;
	case DocumentObjectType::ShapePath:
		return L"ShapePath";
		break;
	case DocumentObjectType::ShapeRect:
		return L"ShapeRect";
		break;
	case DocumentObjectType::SmartTag:
		return L"SmartTag";
		break;
	case DocumentObjectType::SpecialChar:
		return L"SpecialChar";
		break;
	case DocumentObjectType::StructureDocumentTag:
		return L"StructureDocumentTag";
		break;
	case DocumentObjectType::StructureDocumentTagCell:
		return L"StructureDocumentTagCell";
		break;
	case DocumentObjectType::StructureDocumentTagInline:
		return L"StructureDocumentTagInline";
		break;
	case DocumentObjectType::StructureDocumentTagRow:
		return L"StructureDocumentTagRow";
		break;
	case DocumentObjectType::SubDocument:
		return L"SubDocument";
		break;
	case DocumentObjectType::Symbol:
		return L"Symbol";
		break;
	case DocumentObjectType::System:
		return L"System";
		break;
	case DocumentObjectType::Table:
		return L"Table";
		break;
	case DocumentObjectType::TableCell:
		return L"TableCell";
		break;
	case DocumentObjectType::TableRow:
		return L"TableRow";
		break;
	case DocumentObjectType::TextBox:
		return L"TextBox";
		break;
	case DocumentObjectType::TextFormField:
		return L"TextFormField";
		break;
	case DocumentObjectType::TextRange:
		return L"TextRange";
		break;
	case DocumentObjectType::TOC:
		return L"TOC";
		break;
	case DocumentObjectType::XmlParaItem:
		return L"XmlParaItem";
		break;
	case DocumentObjectType::Undefined:
		return L"Undefined";
		break;
	}
	return L"";
}

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Sample.docx";
	wstring outputFile = output_path + L"RecurseAllDocumentObject.txt";

	//Create string builder
	std::wstring builder;

	intrusive_ptr<Document> document =  new Document();
	document->LoadFromFile(inputFile.c_str());

	//find all document object
	int sectionCount = document->GetSections()->GetCount();
	for (int i = 0; i < sectionCount; i++)
	{
		intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
		int SectionIndex = document->GetIndex(section);
		builder.append(L"section index " + to_wstring(SectionIndex) + L" has following ChildObjects");
		builder.append(L"\n");

		int sectionChildObjectsCount = section->GetBody()->GetChildObjects()->GetCount();

		for (int j = 0; j < sectionChildObjectsCount; j++)
		{
			intrusive_ptr<DocumentObject> obj = section->GetBody()->GetChildObjects()->GetItem(j);
			int objIndex = section->GetBody()->GetIndex(obj);
			DocumentObjectType objType = obj->GetDocumentObjectType();

			builder.append(L"Index : " + to_wstring(objIndex) + L", ChildObject Type: " + GetDocumentObjectType(objType));
			builder.append(L"\n");

			if (obj->GetDocumentObjectType() == DocumentObjectType::Paragraph)
			{
				intrusive_ptr<Paragraph> paragraph = Object::Dynamic_cast<Paragraph>(obj);
				int paragraphIndex = section->GetBody()->GetIndex(paragraph);

				builder.append(L"\tParagraph index " + to_wstring(paragraphIndex) + L" has following ChildObjects");
				builder.append(L"\n");

				int paraChildCount = paragraph->GetChildObjects()->GetCount();
				for (int k = 0; k < paraChildCount; k++)
				{
					intrusive_ptr<DocumentObject> obj2 = paragraph->GetChildObjects()->GetItem(k);
					int obj2Index = paragraph->GetIndex(obj2);
					DocumentObjectType obj2Type = obj2->GetDocumentObjectType();

					builder.append(L"\tIndex : " + to_wstring(obj2Index) + L", ChildObject Type: " + GetDocumentObjectType(obj2Type));
					builder.append(L"\n");
				}
			}
		}
		builder.append(L" ");
	}
	//Save to file.
	std::wofstream write(outputFile);
	auto LocUtf8 = locale(locale(""), new std::codecvt_utf8<wchar_t>);
	write.imbue(LocUtf8);
	write << builder;
	write.close();
	document->Close();
}


		