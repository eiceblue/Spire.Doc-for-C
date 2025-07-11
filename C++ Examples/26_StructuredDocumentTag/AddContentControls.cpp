#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddContentControls.docx";

	//Creat a new word document.
	intrusive_ptr<Document> document = new Document();
	intrusive_ptr<Section> section = document->AddSection();
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
	intrusive_ptr<TextRange> txtRange = paragraph->AppendText(L"The following example shows how to add content controls in a Word document.");
	paragraph = section->AddParagraph();

	//Add Combo Box Content Control.
	paragraph = section->AddParagraph();
	txtRange = paragraph->AppendText(L"Add Combo Box Content Control:  ");
	txtRange->GetCharacterFormat()->SetItalic(true);
	intrusive_ptr<StructureDocumentTagInline> sd = new StructureDocumentTagInline(document);
	paragraph->GetChildObjects()->Add(sd);
	sd->GetSDTProperties()->SetSDTType(SdtType::ComboBox);
	intrusive_ptr<SdtComboBox> cb = new SdtComboBox();
	intrusive_ptr<SdtListItem> tempVar = new SdtListItem(L"Spire.Doc");
	cb->GetListItems()->Add(tempVar);
	intrusive_ptr<SdtListItem> tempVar2 = new SdtListItem(L"Spire.XLS");
	cb->GetListItems()->Add(tempVar2);
	intrusive_ptr<SdtListItem> tempVar3 = new SdtListItem(L"Spire.PDF");
	cb->GetListItems()->Add(tempVar3);
	sd->GetSDTProperties()->SetControlProperties(cb);
	intrusive_ptr<TextRange> rt = new TextRange(document);
	rt->SetText(cb->GetListItems()->GetItem(0)->GetDisplayText());
	sd->GetSDTContent()->GetChildObjects()->Add(rt);

	section->AddParagraph();

	//Add Text Content Control.
	paragraph = section->AddParagraph();
	txtRange = paragraph->AppendText(L"Add Text Content Control:  ");
	txtRange->GetCharacterFormat()->SetItalic(true);
	sd = new StructureDocumentTagInline(document);
	paragraph->GetChildObjects()->Add(sd);
	sd->GetSDTProperties()->SetSDTType(SdtType::Text);
	intrusive_ptr<SdtText> text = new SdtText(true);
	text->SetIsMultiline(true);
	sd->GetSDTProperties()->SetControlProperties(text);
	rt = new TextRange(document);
	rt->SetText(L"Text");
	sd->GetSDTContent()->GetChildObjects()->Add(rt);

	section->AddParagraph();

	//Add Picture Content Control.
	paragraph = section->AddParagraph();
	txtRange = paragraph->AppendText(L"Add Picture Content Control:  ");
	txtRange->GetCharacterFormat()->SetItalic(true);
	sd = new StructureDocumentTagInline(document);
	paragraph->GetChildObjects()->Add(sd);
	sd->GetSDTProperties()->SetSDTType(SdtType::Picture);
	intrusive_ptr<DocPicture> pic = new DocPicture(document);
	pic->SetWidth(10);
	pic->SetHeight(10);

	pic->LoadImageSpire(DATAPATH"/logo.png");
	sd->GetSDTContent()->GetChildObjects()->Add(pic);

	section->AddParagraph();

	//Add Date Picker Content Control.
	paragraph = section->AddParagraph();
	txtRange = paragraph->AppendText(L"Add Date Picker Content Control:  ");
	txtRange->GetCharacterFormat()->SetItalic(true);
	sd = new StructureDocumentTagInline(document);
	paragraph->GetChildObjects()->Add(sd);
	sd->GetSDTProperties()->SetSDTType(SdtType::DatePicker);
	intrusive_ptr<SdtDate> date = new SdtDate();
	date->SetCalendarType(CalendarType::Default);
	//date->SetDateFormat(L"yyyy.MM.dd");
	date->SetDateFormatSpire(L"yyyy.MM.dd");
	date->SetFullDate(DateTime::GetNow());
	sd->GetSDTProperties()->SetControlProperties(date);
	rt = new TextRange(document);
	rt->SetText(L"1990.02.08");
	sd->GetSDTContent()->GetChildObjects()->Add(rt);

	section->AddParagraph();

	//Add Drop-Down List Content Control.
	paragraph = section->AddParagraph();
	txtRange = paragraph->AppendText(L"Add Drop-Down List Content Control:  ");
	txtRange->GetCharacterFormat()->SetItalic(true);
	sd = new StructureDocumentTagInline(document);
	paragraph->GetChildObjects()->Add(sd);
	sd->GetSDTProperties()->SetSDTType(SdtType::DropDownList);
	intrusive_ptr<SdtDropDownList> sddl = new SdtDropDownList();
	intrusive_ptr<SdtListItem> tempVar4 = new SdtListItem(L"Harry");
	sddl->GetListItems()->Add(tempVar4);
	intrusive_ptr<SdtListItem> tempVar5 = new SdtListItem(L"Jerry");
	sddl->GetListItems()->Add(tempVar5);
	sd->GetSDTProperties()->SetControlProperties(sddl);
	rt = new TextRange(document);
	rt->SetText(sddl->GetListItems()->GetItem(0)->GetDisplayText());
	sd->GetSDTContent()->GetChildObjects()->Add(rt);

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}