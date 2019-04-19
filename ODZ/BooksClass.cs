using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ODZ
{
    public class BooksClass
    {
        public string IDBook { get; set; }
        public string NameAvtor { get; set; }
        public string Title { get; set; }
        public int Date { get; set; }
        public string Placing { get; set; }

        public BooksClass(string IDBook, string NameAvtor, string Title, int Date, string Placing)
        {
            this.IDBook = IDBook;
            this.NameAvtor = NameAvtor;
            this.Title = Title;
            this.Date = Date;
            this.Placing = Placing;
        }
    }
}
