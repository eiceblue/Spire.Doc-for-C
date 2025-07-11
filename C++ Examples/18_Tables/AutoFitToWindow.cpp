#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/AutoFitToWindow.docx";
	std::wstring inputFile = DATAPATH"/TableSample.docx";
	
	//Create a document
	intrusive_ptr<Document> document = new Document();
	//Load file
	document->LoadFromFile(inputFile.c_str());
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));
	//Automatically fit the table to the active window width
	table->AutoFit(AutoFitBehaviorType::AutoFitToWindow);
	//Save to file and launch it
	document->SaveToFile(outputFile.c_str());
	document->Close();
}
