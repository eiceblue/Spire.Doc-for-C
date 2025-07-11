#include "pch.h"

using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"AcceptOrRejectTrackedChanges.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AcceptOrRejectTrackedChange.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Get the first section and the paragraph we want to accept/reject the changes.
	intrusive_ptr<Section> sec = document->GetSections()->GetItemInSectionCollection(0);
	intrusive_ptr<Paragraph> para = sec->GetParagraphs()->GetItemInParagraphCollection(0);

	//Accept the changes or reject the changes.
	para->GetDocument()->AcceptChanges();
	//para.Document.RejectChanges();

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}