namespace ExcelMvc.Bindings
{
    #region Enumerations

    /// <summary>
    /// Binding mode types
    /// </summary>
    public enum ModeType
    {
        /// <summary>
        /// View fields are updated for its model properties
        /// </summary>
        OneWay,

        /// <summary>
        /// Model properties are update from its view fields
        /// </summary>
        OneWayToSource,

        /// <summary>
        /// Model properties and view fields are exchanged
        /// </summary>
        TwoWay
    }

    #endregion Enumerations
}
