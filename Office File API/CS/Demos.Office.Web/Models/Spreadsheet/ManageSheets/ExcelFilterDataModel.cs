using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets {
    public class ExcelFilterDataModel : ExcelFileModelBase, IValidatableObject {
        public override string OutputName { get; set; } = "Filtered_Data";

        [RegularExpression(@"^(ByCellValues|ByDateAndTextValues|ByCustomDateRange|ByCustomNumericRange|ByCustomText|ByTopValues|Dynamic|ByFontColor|ByFillColor|ByBackgroundColor|SortFilteredRange)$",
            ErrorMessage = "Invalid filter mode. Supported modes: By Cell Values, By Date and Text Values, By Custom Date Range, By Custom Numeric Range, By Custom Text Value, By Top Values, Dynamic, By Font Color, By Fill Color, By Background Color, Sort Filtered Range")]
        public FilterType FilterType { get; set; } = FilterType.ByCellValues;

        internal readonly List<object> FilterTypes = new() {
            new { Text = "By Cell Values", Value = FilterType.ByCellValues },
            new { Text = "By Date And Text Values", Value = FilterType.ByDateAndTextValues },
            new { Text = "By Custom Date Range", Value = FilterType.ByCustomDateRange },
            new { Text = "By Custom Numeric Range", Value = FilterType.ByCustomNumericRange },
            new { Text = "By Custom Text Value", Value = FilterType.ByCustomText },
            new { Text = "By Top Values", Value = FilterType.ByTopValues },
            new { Text = "Dynamic", Value = FilterType.Dynamic },
            new { Text = "By Font Color", Value = FilterType.ByFontColor },
            new { Text = "By Fill Color", Value = FilterType.ByFillColor },
            new { Text = "By Background Color", Value = FilterType.ByBackgroundColor },
            new { Text = "Apply Filter And Sort Data", Value = FilterType.SortFilteredRange },
        };

        internal readonly List<object> ReportedDateValues = new() {
            new { Text = "January, 2014", Value = new DateTime(2014, 1, 1) },
            new { Text = "April, 2014", Value = new DateTime(2014, 4, 1) },
            new { Text = "June, 2014", Value = new DateTime(2014, 6, 1) },

            new { Text = "January, 2015", Value = new DateTime(2015, 1, 1) },
            new { Text = "February, 2015", Value = new DateTime(2015, 2, 1) },
            new { Text = "March, 2015", Value = new DateTime(2015, 3, 1) },
            new { Text = "April, 2015", Value = new DateTime(2015, 4, 1) },
            new { Text = "May, 2015", Value = new DateTime(2015, 5, 1) },
            new { Text = "June, 2015", Value = new DateTime(2015, 6, 1) },
            new { Text = "August, 2015", Value = new DateTime(2015, 8, 1) },
            new { Text = "September, 2015", Value = new DateTime(2015, 9, 1) },

            new { Text = "May, 2022", Value = new DateTime(2022, 5, 1) },
            new { Text = "June, 2022", Value = new DateTime(2022, 6, 1) },
        };
        internal readonly List<string> ColumnNames = new() {
            "Region",
            "Product",
            "Sales",
            "Reported Date"
        };
        public IEnumerable<string> SelectedProductFilterValues { get; set; } = new List<string>{
            "Mozzarella di Giovanni",
            "Gorgonzola Telino"
        };
        public DateTime SelectedReportedDateFilterValue { get; set; } = new DateTime(2015, 1, 1);

        public DateTime ReportedDateStartDate { get; set; } = new DateTime(2014, 1, 14);
        public DateTime ReportedDateEndDate { get; set; } = new DateTime(2022, 6, 15);

        [Range(1250, 13500, ErrorMessage = "The field value must be between 1250 and 13500 ")]
        public int SalesFilterMinValue { get; set; } = 5000;
        [Range(1250, 13500, ErrorMessage = "The field value must be between 1250 and 13500 ")]
        public int SalesFilterMaxValue { get; set; } = 8000;

        internal readonly List<string> ProductCellValues = new() {
            "Camembert Pierrot",
            "Gorgonzola Telino",
            "Mascarpone Fabioli",
            "Mozzarella di Giovanni"
        };

        public SortOrder SortOrder { get; set; } = SortOrder.Ascending;

        internal readonly List<SortOrder> SortOrderOptions = new() {
              SortOrder.Ascending,
              SortOrder.Descending
        };

        // Conditional validation for filter values
        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext) {
            if(FilterType == FilterType.ByCellValues) {
                if(SelectedProductFilterValues == null || !SelectedProductFilterValues.Any()) {
                    yield return new ValidationResult(
                        "Select at least one product",
                        new[] { nameof(SelectedProductFilterValues) });
                }
            }
            if(FilterType == FilterType.ByCustomDateRange) {
                if(ReportedDateStartDate >= ReportedDateEndDate) {
                    yield return new ValidationResult(
                        "End Date must be greater than Start Date",
                        new[] { nameof(ReportedDateStartDate), nameof(ReportedDateEndDate) });
                }
            }
            if(FilterType == FilterType.ByCustomNumericRange) {
                if(SalesFilterMinValue >= SalesFilterMaxValue) {
                    yield return new ValidationResult(
                        "Max Value must be greater than Min Value",
                        new[] { nameof(SalesFilterMinValue), nameof(SalesFilterMaxValue) });
                }
            }
        }
    }

    public enum FilterType {
        ByCellValues,
        ByDateAndTextValues, // By Mixed Values ??
        ByCustomDateRange,
        ByCustomNumericRange,
        ByCustomText,
        ByTopValues,
        Dynamic,
        ByFontColor,
        ByFillColor,
        ByBackgroundColor,
        SortFilteredRange
    }
}
