using System;
namespace DocToExcelNoteParser.Models
{
	public class FootNoteToken
	{
        public FootNoteToken(string? footNoteName, string? footNoteContent)
        {
            FootNoteName = footNoteName;
            FootNoteContent = footNoteContent;
        }

        public string? FootNoteName { get; set; }

		public string? FootNoteContent { get; set; }
	}
}

