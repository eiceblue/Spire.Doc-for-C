#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"PictureCaptionCrossReference.docx";

	//Create word document
	intrusive_ptr<Document> document = new Document();

	//Create a new section
	intrusive_ptr<Section> section = document->AddSection();

	//Add the first paragraph
	intrusive_ptr<Paragraph> firstPara = section->AddParagraph();

	//Add the first picture
	intrusive_ptr<Paragraph> par1 = section->AddParagraph();
	par1->GetFormat()->SetAfterSpacing(10);
	wstring imagePath1 = input_path + L"Spire.Doc.png";
	intrusive_ptr<DocPicture> pic1 = par1->AppendPicture(imagePath1.c_str());
	pic1->SetHeight(120);
	pic1->SetWidth(120);
	//Add caption to the picture
	intrusive_ptr<IParagraph> captionParagraph = pic1->AddCaption(L"Figure", CaptionNumberingFormat::Number, CaptionPosition::BelowItem);
	section->AddParagraph();

	//Add the second picture
	intrusive_ptr<Paragraph> par2 = section->AddParagraph();
	wstring imagePath2 = input_path + L"Word.png";
	intrusive_ptr<DocPicture> pic2 = par2->AppendPicture(imagePath2.c_str());

	pic2->SetHeight(120);
	pic2->SetWidth(120);
	//Add caption to the picture
	captionParagraph = pic2->AddCaption(L"Figure", CaptionNumberingFormat::Number, CaptionPosition::BelowItem);
	section->AddParagraph();

	//Create a bookmark
	wstring bookmarkName = L"Figure_2";
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
	paragraph->AppendBookmarkStart(bookmarkName.c_str());
	paragraph->AppendBookmarkEnd(bookmarkName.c_str());

	//Replace bookmark content
	intrusive_ptr<BookmarksNavigator> navigator = new BookmarksNavigator(document);
	navigator->MoveToBookmark(bookmarkName.c_str());
	intrusive_ptr<TextBodyPart> part = navigator->GetBookmarkContent();
	part->GetBodyItems()->Clear();
	part->GetBodyItems()->Add(captionParagraph);
	navigator->ReplaceBookmarkContent(part);

	//Create cross-reference field to point to bookmark "Figure_2"
	intrusive_ptr<Field> field = new Field(document);
	field->SetType(FieldType::FieldRef);
	field->SetCode(L"REF Figure_2 \\p \\h");
	firstPara->GetChildObjects()->Add(field);
	intrusive_ptr<FieldMark> fieldSeparator = new FieldMark(document, FieldMarkType::FieldSeparator);
	firstPara->GetChildObjects()->Add(fieldSeparator);

	//Set the display text of the field
	intrusive_ptr<TextRange> tr = new TextRange(document);
	tr->SetText(L"Figure 2");
	firstPara->GetChildObjects()->Add(tr);

	intrusive_ptr<FieldMark> fieldEnd = new FieldMark(document, FieldMarkType::FieldEnd);
	firstPara->GetChildObjects()->Add(fieldEnd);

	//Update fields
	document->SetIsUpdateFields(true);

	//Save the Word document
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}
