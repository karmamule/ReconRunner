using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.Common;


namespace ReconRunner
{
    public class DataReaderAdapter : DbDataAdapter
    {
        public int FillFromReader(DataTable dataTable, IDataReader dataReader)
        {
            return this.Fill(dataTable, dataReader);
        }
        protected override RowUpdatedEventArgs CreateRowUpdatedEvent(
        DataRow dataRow,
        IDbCommand command,
        StatementType statementType,
        DataTableMapping tableMapping)
        {
            return null;
        }

        protected override RowUpdatingEventArgs CreateRowUpdatingEvent(
        DataRow dataRow,
        IDbCommand command,
        StatementType statementType,
        DataTableMapping tableMapping)
        {
            return null;
        }

        protected override void OnRowUpdated(
        RowUpdatedEventArgs value)
        {
        }
        protected override void OnRowUpdating(
        RowUpdatingEventArgs value)
        {
        }
    }
}
