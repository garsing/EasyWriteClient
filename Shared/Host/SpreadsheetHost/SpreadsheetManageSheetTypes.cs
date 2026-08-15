using System.Collections.Generic;

namespace WordAddIn1.SpreadsheetHost
{
    internal sealed class SpreadsheetManageSheetRequest
    {
        /// <summary>add | rename | delete</summary>
        public string Action { get; set; }

        /// <summary>add / rename 新名</summary>
        public string Name { get; set; }

        /// <summary>rename / delete 当前名</summary>
        public string Sheet { get; set; }

        /// <summary>add 可选：插在该表之前</summary>
        public string Before { get; set; }

        public bool Confirm { get; set; }
    }

    internal sealed class SpreadsheetManageSheetResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Action { get; set; }

        /// <summary>操作后相关表名（delete 为被删名）</summary>
        public string Sheet { get; set; }

        public List<string> Sheets { get; set; }
    }
}
