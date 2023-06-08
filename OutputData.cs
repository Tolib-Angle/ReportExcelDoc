namespace Izdel
{
    internal class OutputData
    {
        private long _id;
        private string _name;
        private int _kol;
        private decimal _cost;
        private decimal _price;

        public long Id
        {
            get { return _id; }
            set { _id = value; }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        public int Kol
        {
            get { return _kol; }
            set { _kol = value; }
        }

        public decimal Cost
        {
            get { return _cost; }
            set { _cost = value; }
        }

        public decimal Price
        {
            get { return _price; }  
            set { _price = value; }
        }
    }
}
