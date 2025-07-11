#include "pch.h"


using namespace Spire::Doc;

intrusive_ptr<DateTime> LocalTimeToGreenwishTime(intrusive_ptr<DateTime> lacalTime)
{
	intrusive_ptr<TimeZone> localTimeZone = TimeZone::GetCurrentTimeZone();
	intrusive_ptr<TimeSpan> timeSpan = localTimeZone->GetUtcOffset(lacalTime);

	intrusive_ptr<DateTime> greenwishTime = DateTime::op_Subtraction(lacalTime, timeSpan);
	return greenwishTime;
}

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template.docx";
	wstring outputFile = output_path + L"UpdateLastSavedDate.docx";

	intrusive_ptr<Document> document =  new Document();

	//Load the document from disk
	document->LoadFromFile(inputFile.c_str());

	//Update the last saved date
	document->GetBuiltinDocumentProperties()->SetLastSaveDate(LocalTimeToGreenwishTime(DateTime::GetNow()));
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();

}