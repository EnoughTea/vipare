using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Visio;

namespace Vipare {
    // Experimental stuff here, currently unused.

    /// <summary> Extension methods for <see cref="Shape"/>. </summary>
    internal static class ShapeExtensions {
        /// <remarks>Don't forget to call <see cref="Marshal.ReleaseComObject"/> on the acquired cell. </remarks>
        public static Cell PropertyCell(this Shape shape, short propertyRow, short propertyColumn) {
            Contract.Requires(shape != null);

            return shape.PropertyExists(propertyRow)
                ? shape.CellsSRC[(short)VisSectionIndices.visSectionProp, propertyRow, propertyColumn]
                : null;
        }

        public static bool PropertyExists(this Shape shape, short propertyRow) {
            Contract.Requires(shape != null);

            return shape.CellsSRCExists[(short)VisSectionIndices.visSectionProp, propertyRow,
                (short)VisCellIndices.visCustPropsValue, 0] != 0;
        }

        public static string PropertyLabel(this Shape shape, short propertyRow) {
            Contract.Requires(shape != null);

            string result = string.Empty;
            var cell = shape.PropertyCell(propertyRow, (short)VisCellIndices.visCustPropsLabel);
            if (cell != null) {
                result = cell.ResultStr[VisUnitCodes.visNoCast];
                Marshal.ReleaseComObject(cell);
            }

            return result;
        }

        public static string PropertyValue(this Shape shape, short propertyRow) {
            Contract.Requires(shape != null);

            string result = string.Empty;
            var cell = shape.PropertyCell(propertyRow, (short)VisCellIndices.visCustPropsValue);
            if (cell != null) {
                result = cell.ResultStr[VisUnitCodes.visNoCast];
                Marshal.ReleaseComObject(cell);
            }

            return result;
        }

        /// <remarks>Don't forget to call <see cref="Marshal.ReleaseComObject"/> on the acquired selection. </remarks>
        public static Selection SelectAllShapes(this Page page) {
            var selection = page.CreateSelection(VisSelectionTypes.visSelTypeEmpty);
            var shapes = page.Shapes;
            for (int j = 0; j < shapes.Count; j++) {
                Shape shape = shapes[j];
                selection.Select(shape, (short)VisSelectArgs.visSelect);
                Marshal.ReleaseComObject(shape);
            }

            Marshal.ReleaseComObject(shapes);
            return selection;
        }

        /// <summary>
        ///  Recursively traverses a collection of shapes and gathers custom properties for each shape.
        /// </summary>
        /// <param name="shapes">Shapes collection.</param>
        public static IEnumerable<ShapeProperties> GatherShapeProperties(this Shapes shapes) {
            Contract.Requires(shapes != null);

            for (int i = 0; i < shapes.Count; i++) {
                Shape shape = shapes[i];
                var shapeProperties = new ShapeProperties(shape.ID, shape.NameU);
                short propertyRow = (short)VisRowIndices.visRowFirst;
                while (shape.PropertyExists(propertyRow)) {
                    string label = shape.PropertyLabel(propertyRow);
                    string value = shape.PropertyValue(propertyRow);
                    shapeProperties.Properties.Add(new ShapeProperty(propertyRow, label, value));
                    propertyRow++;
                }

                yield return shapeProperties;

                // Print child shapes as well, ignore child shapes for a Master shape (shape dropped from stencil):
                var master = shape.Master;
                var childShapes = shape.Shapes;
                if (master == null && childShapes.Count > 0) {
                    foreach (var childShapeProperties in GatherShapeProperties(childShapes)) {
                        yield return childShapeProperties;
                    }
                } else if (master != null) {
                    Marshal.ReleaseComObject(master);
                }

                Marshal.ReleaseComObject(childShapes);
                Marshal.ReleaseComObject(shape);
            }
        }

        public class ShapeProperties : IEnumerable<ShapeProperty> {
            public ShapeProperties(int id, string name, IEnumerable<ShapeProperty> properties = null) {
                Id = id;
                Name = name;
                Properties = (properties != null) ? new List<ShapeProperty>(properties) : new List<ShapeProperty>();
            }

            public int Id { get; }

            public string Name { get; }

            public List<ShapeProperty> Properties { get; }

            public IEnumerator<ShapeProperty> GetEnumerator() {
                return Properties.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator() {
                return GetEnumerator();
            }
        }

        public struct ShapeProperty {
            public ShapeProperty(short propertyRow, string label, string value) {
                PropertyRow = propertyRow;
                Label = label;
                Value = value;
            }

            public short PropertyRow;
            public string Label;
            public string Value;
        }
    }
}
