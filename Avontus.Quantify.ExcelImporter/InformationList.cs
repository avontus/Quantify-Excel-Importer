using System;
using System.Data;
using System.Data.SqlClient;
using Avontus.Core;
using Avontus.Core.Data;
using Avontus.Rental.Library;

namespace Avontus.Quantify.ExcelImporter
{
    [Serializable()]
    public class InformationList :
        Avontus.Core.ReadOnlyListBase<InformationList, InformationListItem>
    {

#region Declarations

        private int _sortOrder = 0;

#endregion

#region Business Properties

        public int SortOrder
        {
            get { return _sortOrder; }
            set { _sortOrder = value; }
        }

        public bool Contains(string message)
        {
            foreach (InformationListItem item in this)
            {
                if (item.TheMessage.ToUpper() == message.ToUpper())
                    return true;
            }

            return false;
        }

        public bool ContainsErrorsOrFailures()
        {
            foreach (InformationListItem item in this)
            {
                if (item.TheImage.Tag.ToString() == "error" ||
                    item.TheImage.Tag.ToString() == "fail")
                    return true;
            }

            return false;
        }

        public void ResetSortOrder(int NewValue)
        {
            foreach (InformationListItem item in this)
            {
                item.SortOrder = NewValue;
            }
        }

        public void AddSpacer()
        {

            AddSpacer("", SortOrder++);
        }

        public void AddSpacer(string message)
        {

            AddSpacer(message, SortOrder++);
        }

        public void AddSpacer(string message, int sortOrder)
        {
            InformationListItem infoItem = InformationListItem.NewInformationListItem(
                  Properties.Resources.spacer, message);
            infoItem.TheImage.Tag = "spacer";
            infoItem.SortOrder = sortOrder;
            AddItem(infoItem);
        }

        public void AddPassedItem(string message)
        {
            AddPassedItem(message, SortOrder++);
        }

        public void AddPassedItem(string message, int sortOrder)
        {

            InformationListItem infoItem = InformationListItem.NewInformationListItem(
                Properties.Resources.pass, message);

            infoItem.TheImage.Tag = "pass";
            infoItem.SortOrder = sortOrder;
            AddItem(infoItem);
        }

        public void AddWarningItem(string message, int sortOrder)
        {
            
            InformationListItem infoItem = InformationListItem.NewInformationListItem(
                Properties.Resources.warning, message);
            infoItem.TheImage.Tag = "warning";
            infoItem.SortOrder = sortOrder;
            AddItem(infoItem);
        }

        public void AddWarningItem(string message)
        {
            AddWarningItem(message, SortOrder++);
        }

        public void AddInformationItem(string message, int sortOrder)
        {
            InformationListItem infoItem = InformationListItem.NewInformationListItem(
                Properties.Resources.info, message);
            infoItem.TheImage.Tag = "info";
            infoItem.SortOrder = sortOrder;
            AddItem(infoItem);
        }

        public void AddInformationItem(string message)
        {
            AddInformationItem(message, SortOrder++);
        }

        public void AddFailureItem(string message, int sortOrder)
        {
            InformationListItem infoItem = InformationListItem.NewInformationListItem(
                Properties.Resources.fail, message);
            infoItem.TheImage.Tag = "fail";
            infoItem.SortOrder = sortOrder;
            AddItem(infoItem);
        }

        public void AddFailureItem(string message)
        {
            AddFailureItem(message, SortOrder++);
        }

        public void AddItem(InformationListItem infoItem)
        {
            this.IsReadOnly = false;
            this.Add(infoItem);
            this.IsReadOnly = true;
        }

#endregion

#region Factory Methods

        private InformationList()
        { /* require use of factory method */ }

        public static InformationList NewInformationList()
        {
            return new InformationList();
        }
#endregion //Factory Methods

#region Criteria


#endregion

#region Data Access


#endregion

    }
}


