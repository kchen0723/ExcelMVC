namespace ExcelMvc.Runtime
{
    using System;
    using System.Collections.Generic;

    [Serializable]
    internal class TypeResult
    {
        #region Properties

        public Exception Error
        {
            get;
            set;
        }

        public List<string> Types
        {
            get;
            set;
        }

        #endregion Properties
    }
}
