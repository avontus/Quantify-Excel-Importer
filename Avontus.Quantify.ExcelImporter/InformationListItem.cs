using System;
using System.Data.SqlClient;
using System.Data;
using Avontus.Core;
using Avontus.Core.Data;
using System.Drawing;

namespace Avontus.Quantify.ExcelImporter
{

    /// <summary>
    /// This is a lightweight object used for the UI to pass and bind
    /// objects to a list of errors or messages. This object is
    /// exclusively controlled ErrorList.cs.
    /// Note that there is no dataportal objects
    /// </summary>
    [Serializable()]
    public class InformationListItem : Avontus.Core.ReadOnlyBase<InformationListItem>
    {

#region Constructors

        private InformationListItem()
        {
        }

#endregion

#region Declarations

        private Image _theImage;
        private string _message = "";

        // Set sort order
        private int _sortOrder;

#endregion

#region Business Properties and Methods


        public Image TheImage
        {
            [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
            get
            {
                return _theImage;
            }
            [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
            set
            {
               _theImage = value;
            }
        }

        public string TheMessage
        {
            [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
            get
            {
                return _message;
            }
            [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
            set
            {
                _message = value;
            }
        }

        public int SortOrder
        {
            [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
            get
            {
                return _sortOrder;
            }
            [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
            set
            {
                _sortOrder = value;
            }
        }

#endregion

#region Validation Rules

        //protected override void AddBusinessRules()
        //{
        //    //TODO Add validation rules
        //}

#endregion

#region Authorization Rules

        protected override void AddAuthorizationRules()
        {
            // TODO: add authorization rules
            //AuthorizationRules.AllowWrite("", "");
        }

        public static bool CanAddObject()
        {
            // TODO: customize to check user role
            //return ApplicationContext.User.IsInRole("");
            return true;
        }

        public static bool CanGetObject()
        {
            // TODO: customize to check user role
            //return ApplicationContext.User.IsInRole("");
            return true;
        }

        public static bool CanEditObject()
        {
            // TODO: customize to check user role
            //return ApplicationContext.User.IsInRole("");
            return true;
        }

        public static bool CanDeleteObject()
        {
            // TODO: customize to check user role
            //return ApplicationContext.User.IsInRole("");
            return true;
        }

#endregion

#region System.Object Overrides

        protected override object GetIdValue()
        {
            return _message;
        }
        public override string ToString()
        {
            return _message;
        }

#endregion

#region Factory Methods

        public static InformationListItem NewInformationListItem(Image theImage, string theMessage)
        {
            InformationListItem item = new InformationListItem();
            item.TheImage = theImage;
            item.TheMessage = theMessage;
            return item;
        }

#endregion

#region Criteria

#endregion

#region Data Access

#endregion

    }
}