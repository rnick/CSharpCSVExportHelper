using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace rnick.helper.csv.export
{
    /// <summary>
    ///     generates a string containing the CSVs of a list-object.
    /// </summary>
    public class CSVExportHelper
    {
        #region Public Methods and Operators

        /// <summary>
        /// Generates the specified item list.
        /// </summary>
        /// <param name="itemList">
        /// The item list. The list is a list of objects which will parsed for fields or optionally
        ///     properties to get values from.
        /// </param>
        /// <param name="createFieldNamesRow">
        /// if set to <c>true</c>, create a row for FieldNames in CSV-header.
        /// </param>
        /// <param name="delimiter">
        /// Character to use as delimiter in CSV.
        /// </param>
        /// <param name="quoteString">
        /// The character for string quotation.
        /// </param>
        /// <param name="useProperties">
        /// if set to <c>true</c> [use properties instead of fields].
        /// </param>
        /// <param name="valueNames">
        /// The field names to get values from. If useProperties is set to true, properties will be used
        ///     instead of fields. This parameter is optional, if the value is NULL, all found fields / properties of the
        ///     list-objects are used.
        /// </param>
        /// <returns>
        /// csv string of list
        /// </returns>
        public static string Generate(
            List<object> itemList,
            bool createFieldNamesRow,
            string delimiter,
            string quoteString,
            bool useProperties = false,
            List<string> valueNames = null)
        {
            string result = string.Empty;

            // Return empty List, if we have no items
            if (!itemList.Any())
            {
                return string.Empty;
            }

            // Get FieldNames from first list-item, if we have no optional fieldset
            if (valueNames == null || !valueNames.Any())
            {
                valueNames = useProperties
                                 ? itemList[0].GetType()
                                       .GetProperties()
                                       .Select(propertyName => propertyName.Name)
                                       .ToList()
                                 : itemList[0].GetType().GetFields().Select(fieldName => fieldName.Name).ToList();
            }

            // Create title row with field names
            if (createFieldNamesRow)
            {
                result += CreateTitle(delimiter, valueNames);
            }

            // Get all Values from Fields and write to csv
            return itemList.Aggregate(
                result,
                (current, item) => current + CreateRow(delimiter, valueNames, item, quoteString, useProperties));
        }

        #endregion Public Methods and Operators

        #region Methods

        /// <summary>
        /// Creates the row.
        /// </summary>
        /// <param name="delimiter">
        /// The delimiter.
        /// </param>
        /// <param name="fields">
        /// The fields.
        /// </param>
        /// <param name="item">
        /// The item.
        /// </param>
        /// <param name="quoteString">
        /// The quote string.
        /// </param>
        /// <param name="useProperty">
        /// if set to <c>true</c> [use property].
        /// </param>
        /// <returns>
        /// string row of csv data
        /// </returns>
        private static string CreateRow(
            string delimiter,
            IEnumerable<string> fields,
            object item,
            string quoteString,
            bool useProperty)
        {
            string result = string.Empty;
            List<string> row = new List<string>();

            foreach (string field in fields)
            {
                // Check if field or property exists
                bool hasField = useProperty
                                    ? item.GetType().GetProperty(field) == null
                                    : item.GetType().GetField(field) == null;

                if (hasField)
                {
                    // no field / property found
                    row.Add(string.Format("\"No Field \'{0}\' found\"", field));
                }
                else
                {
                    var fieldValue = useProperty
                                         ? item.GetType().GetProperty(field).GetValue(item, null)
                                         : item.GetType().GetField(field).GetValue(item);

                    // check if fieldvalue exist
                    if (fieldValue == null)
                    {
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(quoteString) && (fieldValue is string))
                    {
                        row.Add(string.Format("{0}{1}{0}", quoteString, fieldValue));
                    }
                    else if (fieldValue is decimal)
                    {
                        row.Add(((decimal)fieldValue).ToString(CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        row.Add(fieldValue.ToString());

                        // throw new Exception("QuoteString is null or whitespace");
                    }
                }
            }

            result += string.Join(delimiter, row) + Environment.NewLine;
            return result;
        }

        /// <summary>
        /// Creates the title.
        /// </summary>
        /// <param name="delimiter">
        /// The delimiter.
        /// </param>
        /// <param name="fields">
        /// The fields.
        /// </param>
        /// <returns>
        /// string: csv title row
        /// </returns>
        private static string CreateTitle(string delimiter, IEnumerable<string> fields)
        {
            string result = string.Empty;

            // Create fieldnames line for cvs-title
            List<string> title = fields.ToList();

            result += string.Join(delimiter, title) + Environment.NewLine;
            return result;
        }

        #endregion Methods
    }
}
