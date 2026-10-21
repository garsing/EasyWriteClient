using WordAddIn1.DocumentHost;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class CheckpointKindAdapters
    {
        public static readonly ICheckpointKindAdapter Word = new WordCheckpointKindAdapter();
        public static readonly ICheckpointKindAdapter Workbook = new WorkbookCheckpointKindAdapter();
        public static readonly ICheckpointKindAdapter Presentation = new PresentationCheckpointKindAdapter();

        public static bool TryFromChannel(IOperationChannel channel, out CheckpointLiveTarget target, out string error)
        {
            target = null;
            error = null;
            if (channel == null)
            {
                error = "无可用文档渠道";
                return false;
            }

            string uuid = ChannelRegistry.TryGetChannelDocUuid(channel);
            if (channel is WordChannel word && word.TryGetLiveDocument(out Word.Document wordDoc))
            {
                target = Make(Word, channel, uuid, wordDoc);
                return true;
            }

            if (channel is WpsChannel wps
                && wps.TryGetLiveDocument(out object wpsDoc)
                && DocumentHostCompat.TryAsWordDocument(wpsDoc, out Word.Document compat))
            {
                target = Make(Word, channel, uuid, compat);
                return true;
            }

            if (channel is ExcelChannel excel && excel.TryGetLiveWorkbook(out Excel.Workbook excelBook))
            {
                target = Make(Workbook, channel, uuid, excelBook);
                return true;
            }

            if (channel is EtChannel et && et.TryGetLiveWorkbook(out object etBook))
            {
                target = Make(Workbook, channel, uuid, etBook);
                return true;
            }

            if (channel is PptChannel ppt && ppt.TryGetLivePresentation(out PowerPoint.Presentation pres))
            {
                target = Make(Presentation, channel, uuid, pres);
                return true;
            }

            if (channel is WppChannel wpp && wpp.TryGetLivePresentation(out object wppPres))
            {
                target = Make(Presentation, channel, uuid, wppPres);
                return true;
            }

            error = "当前渠道不是文字稿、工作簿或演示文稿，无法建副本";
            return false;
        }

        private static CheckpointLiveTarget Make(
            ICheckpointKindAdapter adapter,
            IOperationChannel channel,
            string uuid,
            object live)
        {
            return new CheckpointLiveTarget
            {
                Adapter = adapter,
                Channel = channel,
                ChannelId = channel.ChannelId,
                DocUuid = uuid,
                Live = live
            };
        }
    }
}
