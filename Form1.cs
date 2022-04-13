namespace Element_Properties_Window
{
    public partial class Form1 : Form
    {

        private ToolStripMenuItem ContextMerge;

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
            ContextMerge.DropDownItems.Add("Merge Up").Click += MergeUp_Click;
            ContextMerge.DropDownItems.Add("Merge Down").Click += MergeDown_Click;
        }

        private void MergeDown_Click(object? sender, EventArgs e)
        {
            var ParentTD = FindParentTD(htmlEditControl1.CurrentWindowsFormsElement);

            if (ParentTD == null) return;

            var Index = FindIndexInParent(ParentTD.Parent, ParentTD);

            if (Index == -1) return;
            if (ParentTD.Parent.NextSibling is null) return;

            var RowSpan = GetNewRowSpan(ParentTD).ToString();
            HtmlElement RemoveElement = ParentTD.Parent;

            for (int x = 1; x < int.Parse(RowSpan); x++)
            {
                RemoveElement = RemoveElement.NextSibling;
            }

            if (RemoveElement is null) return;

            ((dynamic)RemoveElement.DomElement).removeChild(RemoveElement.Children[Index - 1].DomElement);
            ParentTD.SetAttribute("rowspan", GetNewRowSpan(ParentTD).ToString());

            htmlEditControl1.MoveCursorToElement(ParentTD, Zoople.HTMLEditControl._ELEM_ADJACENCY.ELEM_ADJ_AfterBegin);
        }

        private void MergeUp_Click(object? sender, EventArgs e)
        {
            var ParentTD = FindParentTD(htmlEditControl1.CurrentWindowsFormsElement);

            if (ParentTD == null) return;

            var ColIndex = FindIndexInParent(ParentTD.Parent, ParentTD);
            var RowIndex = FindIndexInParent(ParentTD.Parent.Parent, ParentTD.Parent);

            if (ColIndex == -1 || RowIndex == -1) return;

            HtmlElement NewParentRow = ParentTD.Parent.Parent.Children[RowIndex - 1];
            NewParentRow.Children[ColIndex].SetAttribute("rowspan", GetNewRowSpan(ParentTD).ToString());

            ((dynamic)ParentTD.Parent.DomElement).removeChild(ParentTD.DomElement);

            htmlEditControl1.MoveCursorToElement(ParentTD, Zoople.HTMLEditControl._ELEM_ADJACENCY.ELEM_ADJ_AfterBegin);
        }

        private void MergeRight_Click(object? sender, EventArgs e)
        {
            var ParentTD = FindParentTD(htmlEditControl1.CurrentWindowsFormsElement);

            if (ParentTD == null) return;
            if (ParentTD.NextSibling == null) return;

            ((dynamic)ParentTD.Parent.DomElement).removeChild(ParentTD.NextSibling.DomElement);
            ParentTD.SetAttribute("colspan", GetNewColSpan(ParentTD).ToString());

            htmlEditControl1.MoveCursorToElement(ParentTD, Zoople.HTMLEditControl._ELEM_ADJACENCY.ELEM_ADJ_AfterBegin);
        }

        private void MergeLeft_Click(object? sender, EventArgs e)
        {
            var ParentTD = FindParentTD(htmlEditControl1.CurrentWindowsFormsElement);

            if (ParentTD == null) return;

            var Index = FindIndexInParent(ParentTD.Parent, ParentTD);

            if (Index == -1) return;

            ((dynamic)ParentTD.Parent.DomElement).removeChild(ParentTD.Parent.Children[Index - 1].DomElement);
            ParentTD.SetAttribute("colspan", GetNewColSpan(ParentTD).ToString());

            htmlEditControl1.MoveCursorToElement(ParentTD, Zoople.HTMLEditControl._ELEM_ADJACENCY.ELEM_ADJ_AfterBegin);
        }
        private int GetNewColSpan(HtmlElement Element)
        {
            var ColSpan = Element.GetAttribute("colspan");
            return Int32.Parse(ColSpan) + 1;
        }

        private int GetNewRowSpan(HtmlElement Element)
        {
            var RowSpan = Element.GetAttribute("rowspan");
            return Int32.Parse(RowSpan) + 1;
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