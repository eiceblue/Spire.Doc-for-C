#include "pch.h"
using namespace Spire::Doc;

void SetPage(intrusive_ptr<Section> section);
void InsertCover(intrusive_ptr<Section> section);
void InsertContent(intrusive_ptr<Section> section);
wstring GetFieldType(FieldType type);
bool IsFieldType(FieldType type);
int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertBreak.docx";
	//Create word document
	intrusive_ptr<Document> document = new Document();

	intrusive_ptr<Section> section = document->AddSection();

	//page setup
	SetPage(section);

	//Add cover.
	InsertCover(section);

	//insert a break code
	section = document->AddSection();
	section->AddParagraph()->InsertSectionBreak(SectionBreakType::NewPage);

	//add content
	InsertContent(section);

	//Save as doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();

}

void SetPage(intrusive_ptr<Section> section)
{
	//the unit of all measures below is point, 1point = 0.3528 mm
	section->GetPageSetup()->SetPageSize(PageSize::A4());
	section->GetPageSetup()->GetMargins()->SetTop(72.0f);
	section->GetPageSetup()->GetMargins()->SetBottom(72.0f);
	section->GetPageSetup()->GetMargins()->SetLeft(89.85f);
	section->GetPageSetup()->GetMargins()->SetRight(89.85f);
}

void InsertCover(intrusive_ptr<Section> section)
{
	intrusive_ptr<ParagraphStyle> smallStyle = new ParagraphStyle(section->GetDocument());
	smallStyle->SetName(L"small");
	smallStyle->GetCharacterFormat()->SetFontName(L"Arial");
	smallStyle->GetCharacterFormat()->SetFontSize(9);
	smallStyle->GetCharacterFormat()->SetTextColor(Color::GetGray());
	section->GetDocument()->GetStyles()->Add(smallStyle);

	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
	paragraph->AppendText(L"The sample demonstrates how to insert section break.");
	paragraph->ApplyStyle(smallStyle->GetName());

	intrusive_ptr<Paragraph> title = section->AddParagraph();
	intrusive_ptr<TextRange> text = title->AppendText(L"Field Types Supported by Spire.Doc");
	text->GetCharacterFormat()->SetFontName(L"Arial");
	text->GetCharacterFormat()->SetFontSize(20);
	text->GetCharacterFormat()->SetBold(true);
	title->GetFormat()->SetBeforeSpacing(section->GetPageSetup()->GetPageSize()->GetHeight() / 2 - 3 * section->GetPageSetup()->GetMargins()->GetTop());
	title->GetFormat()->SetAfterSpacing(8);
	title->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"e-iceblue Spire.Doc team.");
	paragraph->ApplyStyle(smallStyle->GetName());
	paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

}

void InsertContent(intrusive_ptr<Section> section)
{
	intrusive_ptr<ParagraphStyle> list = new ParagraphStyle(section->GetDocument());
	list->SetName(L"list");
	list->GetCharacterFormat()->SetFontName(L"Arial");
	list->GetCharacterFormat()->SetFontSize(11);
	list->GetParagraphFormat()->SetLineSpacing(1.5F * 12.0F);
	list->GetParagraphFormat()->SetLineSpacingRule(LineSpacingRule::Multiple);
	section->GetDocument()->GetStyles()->Add(list);

	intrusive_ptr<Paragraph> title = section->AddParagraph();
	intrusive_ptr<TextRange> text = title->AppendText(L"Field type list:");
	title->ApplyStyle(list->GetName());

	bool first = true;

	for (int i = (int)FieldType::FieldNone; i < (int)FieldType::FieldBibliography; i++)
	{
		FieldType type = static_cast<FieldType>(i);
		if (type == FieldType::FieldUnknown || type == FieldType::FieldNone || type == FieldType::FieldEmpty || !IsFieldType(type))
		{
			continue;
		}
		intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
		//paragraph->AppendText(StringHelper::formatSimple(L"{0} is supported in Spire.Doc", type));
		wstring* builder = new wstring();
		wstring targetStr = builder->append(GetFieldType(type)).append(L" is supported in Spire.Doc");
		paragraph->AppendText(targetStr.c_str());
		if (first)
		{
			paragraph->GetListFormat()->ApplyNumberedStyle();
			first = false;
		}
		else
		{
			paragraph->GetListFormat()->ContinueListNumbering();
		}
		paragraph->ApplyStyle(list->GetName());
		delete builder;
	}
}
wstring GetFieldType(FieldType type)
{
	switch (type)
	{
	case Spire::Doc::FieldType::FieldNone:
		return L"FieldNone";
		break;
	case Spire::Doc::FieldType::FieldAddin:
		return L"FieldAddin";
		break;
	case Spire::Doc::FieldType::FieldAdvance:
		return L"FieldAdvance";
		break;
	case Spire::Doc::FieldType::FieldAsk:
		return L"FieldAsk";
		break;
	case Spire::Doc::FieldType::FieldAuthor:
		return L"FieldAuthor";
		break;
	case Spire::Doc::FieldType::FieldAutoNum:
		return L"FieldAutoNum";
		break;
	case Spire::Doc::FieldType::FieldAutoNumLegal:
		return L"FieldAutoNumLegal";
		break;
	case Spire::Doc::FieldType::FieldAutoNumOutline:
		return L"FieldAutoNumOutline";
		break;
	case Spire::Doc::FieldType::FieldAutoText:
		return L"FieldAutoText";
		break;
	case Spire::Doc::FieldType::FieldAutoTextList:
		return L"FieldAutoTextList";
		break;
	case Spire::Doc::FieldType::FieldBarcode:
		return L"FieldBarcode";
		break;
	case Spire::Doc::FieldType::FieldCitation:
		return L"FieldCitation";
		break;
	case Spire::Doc::FieldType::FieldComments:
		return L"FieldComments";
		break;
	case Spire::Doc::FieldType::FieldCompare:
		return L"FieldCompare";
		break;
	case Spire::Doc::FieldType::FieldCreateDate:
		return L"FieldCreateDate";
		break;
	case Spire::Doc::FieldType::FieldData:
		return L"FieldData";
		break;
	case Spire::Doc::FieldType::FieldDatabase:
		return L"FieldDatabase";
		break;
	case Spire::Doc::FieldType::FieldDate:
		return L"FieldDate";
		break;
	case Spire::Doc::FieldType::FieldDDE:
		return L"FieldDDE";
		break;
	case Spire::Doc::FieldType::FieldDDEAuto:
		return L"FieldDDEAuto";
		break;
	case Spire::Doc::FieldType::FieldDocProperty:
		return L"FieldDocProperty";
		break;
	case Spire::Doc::FieldType::FieldDocVariable:
		return L"FieldDocVariable";
		break;
	case Spire::Doc::FieldType::FieldEditTime:
		return L"FieldEditTime";
		break;
	case Spire::Doc::FieldType::FieldEmbed:
		return L"FieldEmbed";
		break;
	case Spire::Doc::FieldType::FieldEmpty:
		return L"FieldEmpty";
		break;
	case Spire::Doc::FieldType::FieldFormula:
		return L"FieldFormula";
		break;
	case Spire::Doc::FieldType::FieldFileName:
		return L"FieldFileName";
		break;
	case Spire::Doc::FieldType::FieldFileSize:
		return L"FieldFileSize";
		break;
	case Spire::Doc::FieldType::FieldFillIn:
		return L"FieldFillIn";
		break;
	case Spire::Doc::FieldType::FieldFootnoteRef:
		return L"FieldFootnoteRef";
		break;
	case Spire::Doc::FieldType::FieldFormCheckBox:
		return L"FieldFormCheckBox";
		break;
	case Spire::Doc::FieldType::FieldFormDropDown:
		return L"FieldFormDropDown";
		break;
	case Spire::Doc::FieldType::FieldFormTextInput:
		return L"FieldFormTextInput";
		break;
	case Spire::Doc::FieldType::FieldEquation:
		return L"FieldEquation";
		break;
	case Spire::Doc::FieldType::FieldGlossary:
		return L"FieldGlossary";
		break;
	case Spire::Doc::FieldType::FieldGoToButton:
		return L"FieldGoToButton";
		break;
	case Spire::Doc::FieldType::FieldHTMLActiveX:
		return L"FieldHTMLActiveX";
		break;
	case Spire::Doc::FieldType::FieldHyperlink:
		return L"FieldHyperlink";
		break;
	case Spire::Doc::FieldType::FieldIf:
		return L"FieldIf";
		break;
	case Spire::Doc::FieldType::FieldImport:
		return L"FieldImport";
		break;
	case Spire::Doc::FieldType::FieldInclude:
		return L"FieldInclude";
		break;
	case Spire::Doc::FieldType::FieldIncludePicture:
		return L"FieldIncludePicture";
		break;
	case Spire::Doc::FieldType::FieldIncludeText:
		return L"FieldIncludeText";
		break;
	case Spire::Doc::FieldType::FieldIndex:
		return L"FieldIndex";
		break;
	case Spire::Doc::FieldType::FieldIndexEntry:
		return L"FieldIndexEntry";
		break;
	case Spire::Doc::FieldType::FieldInfo:
		return L"FieldInfo";
		break;
	case Spire::Doc::FieldType::FieldKeyWord:
		return L"FieldKeyWord";
		break;
	case Spire::Doc::FieldType::FieldLastSavedBy:
		return L"FieldLastSavedBy";
		break;
	case Spire::Doc::FieldType::FieldLink:
		return L"FieldLink";
		break;
	case Spire::Doc::FieldType::FieldListNum:
		return L"FieldListNum";
		break;
	case Spire::Doc::FieldType::FieldMacroButton:
		return L"FieldMacroButton";
		break;
	case Spire::Doc::FieldType::FieldMergeField:
		return L"FieldMergeField";
		break;
	case Spire::Doc::FieldType::FieldMergeRec:
		return L"FieldMergeRec";
		break;
	case Spire::Doc::FieldType::FieldMergeSeq:
		return L"FieldMergeSeq";
		break;
	case Spire::Doc::FieldType::FieldNext:
		return L"FieldNext";
		break;
	case Spire::Doc::FieldType::FieldNextIf:
		return L"FieldNextIf";
		break;
	case Spire::Doc::FieldType::FieldNoteRef:
		return L"FieldNoteRef";
		break;
	case Spire::Doc::FieldType::FieldNumChars:
		return L"FieldNumChars";
		break;
	case Spire::Doc::FieldType::FieldNumPages:
		return L"FieldNumPages";
		break;
	case Spire::Doc::FieldType::FieldNumWords:
		return L"FieldNumWords";
		break;
	case Spire::Doc::FieldType::FieldOCX:
		return L"FieldOCX";
		break;
	case Spire::Doc::FieldType::FieldPage:
		return L"FieldPage";
		break;
	case Spire::Doc::FieldType::FieldPageRef:
		return L"FieldPageRef";
		break;
	case Spire::Doc::FieldType::FieldPrint:
		return L"FieldPrint";
		break;
	case Spire::Doc::FieldType::FieldPrintDate:
		return L"FieldPrintDate";
		break;
	case Spire::Doc::FieldType::FieldPrivate:
		return L"FieldPrivate";
		break;
	case Spire::Doc::FieldType::FieldQuote:
		return L"FieldQuote";
		break;
	case Spire::Doc::FieldType::FieldRef:
		return L"FieldRef";
		break;
	case Spire::Doc::FieldType::FieldRefDoc:
		return L"FieldRefDoc";
		break;
	case Spire::Doc::FieldType::FieldRevisionNum:
		return L"FieldRevisionNum";
		break;
	case Spire::Doc::FieldType::FieldSaveDate:
		return L"FieldSaveDate";
		break;
	case Spire::Doc::FieldType::FieldSection:
		return L"FieldSection";
		break;
	case Spire::Doc::FieldType::FieldSectionPages:
		return L"FieldSectionPages";
		break;
	case Spire::Doc::FieldType::FieldSequence:
		return L"FieldSequence";
		break;
	case Spire::Doc::FieldType::FieldSet:
		return L"FieldSet";
		break;
	case Spire::Doc::FieldType::FieldSkipIf:
		return L"FieldSkipIf";
		break;
	case Spire::Doc::FieldType::FieldStyleRef:
		return L"FieldStyleRef";
		break;
	case Spire::Doc::FieldType::FieldSubject:
		return L"FieldSubject";
		break;
	case Spire::Doc::FieldType::FieldSubscriber:
		return L"FieldSubscriber";
		break;
	case Spire::Doc::FieldType::FieldSymbol:
		return L"FieldSymbol";
		break;
	case Spire::Doc::FieldType::FieldTemplate:
		return L"FieldTemplate";
		break;
	case Spire::Doc::FieldType::FieldTime:
		return L"FieldTime";
		break;
	case Spire::Doc::FieldType::FieldTitle:
		return L"FieldTitle";
		break;
	case Spire::Doc::FieldType::FieldTOA:
		return L"FieldTOA";
		break;
	case Spire::Doc::FieldType::FieldTOAEntry:
		return L"FieldTOAEntry";
		break;
	case Spire::Doc::FieldType::FieldTOC:
		return L"FieldTOC";
		break;
	case Spire::Doc::FieldType::FieldTOCEntry:
		return L"FieldTOCEntry";
		break;
	case Spire::Doc::FieldType::FieldUserAddress:
		return L"FieldUserAddress";
		break;
	case Spire::Doc::FieldType::FieldUserInitials:
		return L"FieldUserInitials";
		break;
	case Spire::Doc::FieldType::FieldUserName:
		return L"FieldUserName";
		break;
	case Spire::Doc::FieldType::FieldShape:
		return L"FieldShape";
		break;
	case Spire::Doc::FieldType::FieldBidiOutline:
		return L"FieldBidiOutline";
		break;
	case Spire::Doc::FieldType::FieldAddressBlock:
		return L"FieldAddressBlock";
		break;
	case Spire::Doc::FieldType::FieldUnknown:
		return L"FieldUnknown";
		break;
	case Spire::Doc::FieldType::FieldCannotParse:
		return L"FieldCannotParse";
		break;
	case Spire::Doc::FieldType::FieldGreetingLine:
		return L"FieldGreetingLine";
		break;
	case Spire::Doc::FieldType::FieldRefNoKeyword:
		return L"FieldRefNoKeyword";
		break;
	case Spire::Doc::FieldType::FieldMacro:
		return L"FieldMacro";
		break;
	case Spire::Doc::FieldType::FieldMergeBarcode:
		return L"FieldMergeBarcode";
		break;
	case Spire::Doc::FieldType::FieldDisplayBarcode:
		return L"FieldDisplayBarcode";
		break;
	case Spire::Doc::FieldType::FieldBibliography:
		return L"FieldBibliography";
		break;
	}
}
bool IsFieldType(FieldType type)
{
	switch (type)
	{
	case Spire::Doc::FieldType::FieldNone:
		return true;
		break;
	case Spire::Doc::FieldType::FieldAddin:
		return true;
		break;
	case Spire::Doc::FieldType::FieldAdvance:
		return true;
		break;
	case Spire::Doc::FieldType::FieldAsk:
		return true;
		break;
	case Spire::Doc::FieldType::FieldAuthor:
		return true;
		break;
	case Spire::Doc::FieldType::FieldAutoNum:
		return true;
		break;
	case Spire::Doc::FieldType::FieldAutoNumLegal:
		return true;
		break;
	case Spire::Doc::FieldType::FieldAutoNumOutline:
		return true;
		break;
	case Spire::Doc::FieldType::FieldAutoText:
		return true;
		break;
	case Spire::Doc::FieldType::FieldAutoTextList:
		return true;
		break;
	case Spire::Doc::FieldType::FieldBarcode:
		return true;
		break;
	case Spire::Doc::FieldType::FieldCitation:
		return true;
		break;
	case Spire::Doc::FieldType::FieldComments:
		return true;
		break;
	case Spire::Doc::FieldType::FieldCompare:
		return true;
		break;
	case Spire::Doc::FieldType::FieldCreateDate:
		return true;
		break;
	case Spire::Doc::FieldType::FieldData:
		return true;
		break;
	case Spire::Doc::FieldType::FieldDatabase:
		return true;
		break;
	case Spire::Doc::FieldType::FieldDate:
		return true;
		break;
	case Spire::Doc::FieldType::FieldDDE:
		return true;
		break;
	case Spire::Doc::FieldType::FieldDDEAuto:
		return true;
		break;
	case Spire::Doc::FieldType::FieldDocProperty:
		return true;
		break;
	case Spire::Doc::FieldType::FieldDocVariable:
		return true;
		break;
	case Spire::Doc::FieldType::FieldEditTime:
		return true;
		break;
	case Spire::Doc::FieldType::FieldEmbed:
		return true;
		break;
	case Spire::Doc::FieldType::FieldEmpty:
		return true;
		break;
	case Spire::Doc::FieldType::FieldFormula:
		return true;
		break;
	case Spire::Doc::FieldType::FieldFileName:
		return true;
		break;
	case Spire::Doc::FieldType::FieldFileSize:
		return true;
		break;
	case Spire::Doc::FieldType::FieldFillIn:
		return true;
		break;
	case Spire::Doc::FieldType::FieldFootnoteRef:
		return true;
		break;
	case Spire::Doc::FieldType::FieldFormCheckBox:
		return true;
		break;
	case Spire::Doc::FieldType::FieldFormDropDown:
		return true;
		break;
	case Spire::Doc::FieldType::FieldFormTextInput:
		return true;
		break;
	case Spire::Doc::FieldType::FieldEquation:
		return true;
		break;
	case Spire::Doc::FieldType::FieldGlossary:
		return true;
		break;
	case Spire::Doc::FieldType::FieldGoToButton:
		return true;
		break;
	case Spire::Doc::FieldType::FieldHTMLActiveX:
		return true;
		break;
	case Spire::Doc::FieldType::FieldHyperlink:
		return true;
		break;
	case Spire::Doc::FieldType::FieldIf:
		return true;
		break;
	case Spire::Doc::FieldType::FieldImport:
		return true;
		break;
	case Spire::Doc::FieldType::FieldInclude:
		return true;
		break;
	case Spire::Doc::FieldType::FieldIncludePicture:
		return true;
		break;
	case Spire::Doc::FieldType::FieldIncludeText:
		return true;
		break;
	case Spire::Doc::FieldType::FieldIndex:
		return true;
		break;
	case Spire::Doc::FieldType::FieldIndexEntry:
		return true;
		break;
	case Spire::Doc::FieldType::FieldInfo:
		return true;
		break;
	case Spire::Doc::FieldType::FieldKeyWord:
		return true;
		break;
	case Spire::Doc::FieldType::FieldLastSavedBy:
		return true;
		break;
	case Spire::Doc::FieldType::FieldLink:
		return true;
		break;
	case Spire::Doc::FieldType::FieldListNum:
		return true;
		break;
	case Spire::Doc::FieldType::FieldMacroButton:
		return true;
		break;
	case Spire::Doc::FieldType::FieldMergeField:
		return true;
		break;
	case Spire::Doc::FieldType::FieldMergeRec:
		return true;
		break;
	case Spire::Doc::FieldType::FieldMergeSeq:
		return true;
		break;
	case Spire::Doc::FieldType::FieldNext:
		return true;
		break;
	case Spire::Doc::FieldType::FieldNextIf:
		return true;
		break;
	case Spire::Doc::FieldType::FieldNoteRef:
		return true;
		break;
	case Spire::Doc::FieldType::FieldNumChars:
		return true;
		break;
	case Spire::Doc::FieldType::FieldNumPages:
		return true;
		break;
	case Spire::Doc::FieldType::FieldNumWords:
		return true;
		break;
	case Spire::Doc::FieldType::FieldOCX:
		return true;
		break;
	case Spire::Doc::FieldType::FieldPage:
		return true;
		break;
	case Spire::Doc::FieldType::FieldPageRef:
		return true;
		break;
	case Spire::Doc::FieldType::FieldPrint:
		return true;
		break;
	case Spire::Doc::FieldType::FieldPrintDate:
		return true;
		break;
	case Spire::Doc::FieldType::FieldPrivate:
		return true;
		break;
	case Spire::Doc::FieldType::FieldQuote:
		return true;
		break;
	case Spire::Doc::FieldType::FieldRef:
		return true;
		break;
	case Spire::Doc::FieldType::FieldRefDoc:
		return true;
		break;
	case Spire::Doc::FieldType::FieldRevisionNum:
		return true;
		break;
	case Spire::Doc::FieldType::FieldSaveDate:
		return true;
		break;
	case Spire::Doc::FieldType::FieldSection:
		return true;
		break;
	case Spire::Doc::FieldType::FieldSectionPages:
		return true;
		break;
	case Spire::Doc::FieldType::FieldSequence:
		return true;
		break;
	case Spire::Doc::FieldType::FieldSet:
		return true;
		break;
	case Spire::Doc::FieldType::FieldSkipIf:
		return true;
		break;
	case Spire::Doc::FieldType::FieldStyleRef:
		return true;
		break;
	case Spire::Doc::FieldType::FieldSubject:
		return true;
		break;
	case Spire::Doc::FieldType::FieldSubscriber:
		return true;
		break;
	case Spire::Doc::FieldType::FieldSymbol:
		return true;
		break;
	case Spire::Doc::FieldType::FieldTemplate:
		return true;
		break;
	case Spire::Doc::FieldType::FieldTime:
		return true;
		break;
	case Spire::Doc::FieldType::FieldTitle:
		return true;
		break;
	case Spire::Doc::FieldType::FieldTOA:
		return true;
		break;
	case Spire::Doc::FieldType::FieldTOAEntry:
		return true;
		break;
	case Spire::Doc::FieldType::FieldTOC:
		return true;
		break;
	case Spire::Doc::FieldType::FieldTOCEntry:
		return true;
		break;
	case Spire::Doc::FieldType::FieldUserAddress:
		return true;
		break;
	case Spire::Doc::FieldType::FieldUserInitials:
		return true;
		break;
	case Spire::Doc::FieldType::FieldUserName:
		return true;
		break;
	case Spire::Doc::FieldType::FieldShape:
		return true;
		break;
	case Spire::Doc::FieldType::FieldBidiOutline:
		return true;
		break;
	case Spire::Doc::FieldType::FieldAddressBlock:
		return true;
		break;
	case Spire::Doc::FieldType::FieldUnknown:
		return true;
		break;
	case Spire::Doc::FieldType::FieldCannotParse:
		return true;
		break;
	case Spire::Doc::FieldType::FieldGreetingLine:
		return true;
		break;
	case Spire::Doc::FieldType::FieldRefNoKeyword:
		return true;
		break;
	case Spire::Doc::FieldType::FieldMacro:
		return true;
		break;
	case Spire::Doc::FieldType::FieldMergeBarcode:
		return true;
		break;
	case Spire::Doc::FieldType::FieldDisplayBarcode:
		return true;
		break;
	case Spire::Doc::FieldType::FieldBibliography:
		return true;
		break;
	default:
		return false;
		break;
	}
}
