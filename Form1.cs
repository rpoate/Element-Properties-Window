namespace Element_Properties_Window
{
    public partial class Form1 : Form
    {

        ToolStripMenuItem ContextMerge;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            //var PropertyGrid = this.htmlEditControl1.Controls.Find("PropertyGrid1", true)[0];
            //PropertyGrid.Width = 600;

            ContextMerge = (ToolStripMenuItem)this.htmlEditControl1.ContextMenuWYSIWYG.Items.Add("Merge Cells");

            ContextMerge.DropDownItems.Add("Merge Left").Click += MergeLeft_Click;
            ContextMerge.DropDownItems.Add("Merge Right").Click += MergeRight_Click;

        }

        private void MergeRight_Click(object? sender, EventArgs e)
        {
            var ParentTD = FindParentTD(htmlEditControl1.CurrentWindowsFormsElement);

            if (ParentTD == null) return;

            ((dynamic)ParentTD.Parent.DomElement).removeChild(ParentTD.NextSibling.DomElement);
            ParentTD.SetAttribute("colspan", "2");
        }

        private void MergeLeft_Click(object? sender, EventArgs e)
        {
            var ParentTD = FindParentTD(htmlEditControl1.CurrentWindowsFormsElement);

            if (ParentTD == null) return;

            var CurrentInnerHTML = ParentTD.InnerHtml;
            var Index = FindIndexInParent(ParentTD.Parent, ParentTD);

            if (Index == -1) return;

            ((dynamic)ParentTD.Parent.DomElement).removeChild(ParentTD.Parent.Children[Index - 1].DomElement);
            ParentTD.SetAttribute("colspan", GetNewColSpan(ParentTD).ToString());

        }

        private int GetNewColSpan(HtmlElement Element)
        {

            var ColSpan = Element.GetAttribute("colspan");

            return Int32.Parse(ColSpan) + 1;
        }

        private int FindIndexInParent(HtmlElement row, HtmlElement cell)
        {
            int count = 0;

            foreach (HtmlElement currentcell in row.Children)
            {
                if (currentcell == cell) return count;
                count++;
            }

            return -1;
        }

        private HtmlElement? FindParentTD(HtmlElement Element)
        {

            if (Element.TagName == "BODY") return null;

            if (Element.TagName == "TD")
            {
                return Element;
            }
            else
            {
                return FindParentTD(Element.Parent);
            }
        }
    }
}