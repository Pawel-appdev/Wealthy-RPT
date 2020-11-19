using System.Data;
using System.Reflection;
using ADODB;


namespace Wealthy_RPT
{
    public static class DTtoRSconvert
    {
        public static Recordset ConvertToRecordSet(DataTable inTable)
        {
            var recordSet = new Recordset { CursorLocation = CursorLocationEnum.adUseClient };

            var recordSetFields = recordSet.Fields;
            var inColumns = inTable.Columns;

            foreach (DataColumn column in inColumns)
            {
                recordSetFields.Append(column.ColumnName
                                    , TranslateType(column.DataType)
                                    , column.MaxLength
                                    , column.AllowDBNull
                                          ? FieldAttributeEnum.adFldIsNullable
                                          : FieldAttributeEnum.adFldUnspecified
                                    , null);
            }

            recordSet.Open(Missing.Value
                        , Missing.Value
                        , CursorTypeEnum.adOpenStatic
                        , LockTypeEnum.adLockOptimistic, 0);

            foreach (DataRow row in inTable.Rows)
            {
                recordSet.AddNew(Missing.Value,
                              Missing.Value);

                for (var columnIndex = 0; columnIndex < inColumns.Count; columnIndex++)
                {
                    recordSetFields[columnIndex].Value = row[columnIndex];
                }
            }

            return recordSet;
        }

        private static DataTypeEnum TranslateType(IReflect columnDataType)
        {
            switch (columnDataType.UnderlyingSystemType.ToString())
            {
                case "System.Boolean":
                    return DataTypeEnum.adBoolean;

                case "System.Byte":
                    return DataTypeEnum.adUnsignedTinyInt;

                case "System.Char":
                    return DataTypeEnum.adChar;

                case "System.DateTime":
                    return DataTypeEnum.adDate;

                case "System.Decimal":
                    return DataTypeEnum.adCurrency;

                case "System.Double":
                    return DataTypeEnum.adDouble;

                case "System.Int16":
                    return DataTypeEnum.adSmallInt;

                case "System.Int32":
                    return DataTypeEnum.adInteger;

                case "System.Int64":
                    return DataTypeEnum.adBigInt;

                case "System.SByte":
                    return DataTypeEnum.adTinyInt;

                case "System.Single":
                    return DataTypeEnum.adSingle;

                case "System.UInt16":
                    return DataTypeEnum.adUnsignedSmallInt;

                case "System.UInt32":
                    return DataTypeEnum.adUnsignedInt;

                case "System.UInt64":
                    return DataTypeEnum.adUnsignedBigInt;

                //System.String will also be handled by default
                default:
                    return DataTypeEnum.adVarChar;
            }
        }
    }
}
