import datetime as dt
from dataclasses import dataclass


@dataclass
class TimeStruct:
    dateTimeStart: dt.datetime
    dateTimeEnd: dt.datetime
    sheetName: str
    fullSheetName: str

    def findUnusedSheetName(self, baseName, xlsxWriter):
        replacements = [',', '@', ';', '!', '#', '^', '+', '~']

        for replacement in replacements:
            candidate = f"{baseName} {self.sheetName.replace('*', replacement)}"
            if candidate not in xlsxWriter.sheets:
                return candidate

    @staticmethod
    def normalizeDict(dictt, df):
        ts = dictt['time_start']
        te = dictt['time_end']

        dictt['dateTime_start'] = df['Capture_time'].min().strftime('%d.%m.%Y') + " - " + ts
        dictt['dateTime_end'] = df['Capture_time'].max().strftime('%d.%m.%Y') + " - " + te

    @staticmethod
    def createFromStartAndEndTime(startTime, endTime, df):
        dictt = {'time_start': startTime, 'time_end': endTime}
        return TimeStruct.createFromDict(dictt, df)

    @staticmethod
    def createFromDict(dictt, df):
        if 'dateTime_start' not in dictt:
            TimeStruct.normalizeDict(dictt, df)

        dts = dt.datetime.strptime(dictt['dateTime_start'], "%d.%m.%Y - %H:%M")
        dte = dt.datetime.strptime(dictt['dateTime_end'], "%d.%m.%Y - %H:%M")

        name = f'{dts.strftime("%d.%m.%y* %H.%M")}'
        fullName = f'{dts.strftime("%d.%m.%Y - %H:%M")} -> {dte.strftime("%d.%m.%Y - %H:%M")}'
        return TimeStruct(dts, dte, name, fullName)
