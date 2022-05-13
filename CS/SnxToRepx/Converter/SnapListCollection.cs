using DevExpress.Snap.Core.API;
using DevExpress.XtraRichEdit.API.Native;
using System.Collections.Generic;

namespace SnxToRepx.Converter
{
    /// <summary>
    /// This is a collection of SnapList fields with additional logic
    /// </summary>
    class SnapListCollection : List<Field>
    {
        private SnapDocument source; //The source document
        public SnapListCollection(SnapDocument source)
        {
            this.source = source;
        }

        /// <summary>
        /// This method collects top-level SnapLists
        /// </summary>
        public void PrepareCollection()
        {
            //Clear the collection
            this.Clear();

            //Iterate through fields
            foreach (Field field in source.Fields)
            {
                //Get a top-level field for the current field
                Field topField = GetTopLevelField(field);
                if (!this.ContainsField(topField))
                {
                    //If the field is not yet added and it is a SnapList, add it to the collection
                    SnapEntity entity = source.ParseField(topField);
                    if (entity != null && entity is SnapList)
                        this.Add(topField);
                }
            }

            //Sort lists based on their ranges (SnapLists can be in the reverse order in the document model)
            this.Sort(new FieldComparer());
        }

        /// <summary>
        /// Gets a SnapList instance based on its index in the collection
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public SnapList GetSnapList(int index)
        {
            //Get and parse the source field
            Field field = this[index];
            SnapEntity entity = source.ParseField(field);
            //Return SnapList
            return (SnapList)entity;
        }

        /// <summary>
        /// Gets a SnapList instance containing the specified position
        /// </summary>
        /// <param name="position">A position to check</param>
        /// <returns>The SnapList</returns>
        public SnapList GetSnapListByPosition(int position)
        {
            //Iterate through elements in the collection
            for (int i = 0; i < this.Count; i++)
            {
                //Get a field and compare its range with the specified position
                Field field = this[i];
                if (field.Range.Start.ToInt() <= position &&
                    field.Range.End.ToInt() >= position)
                    //If they match, return SnapList
                    return (SnapList)source.ParseField(field);
            }
            //If no SnapList is found, return Null
            return null;
        }

        /// <summary>
        /// Gets the top-level field for the current field
        /// </summary>
        /// <param name="field">The currently processed field</param>
        /// <returns>The top-level field</returns>
        private Field GetTopLevelField(Field field)
        {
            //Access the parent field
            Field parentField = field;
            while (parentField.Parent != null)
                parentField = parentField.Parent;

            //Return the resulting field
            return parentField;
        }   
        
        /// <summary>
        /// Checks whether the collection contains a specific field
        /// </summary>
        /// <param name="field">A field to locate</param>
        /// <returns>True if the collection contains the field. Otherwise, false.</returns>
        public bool ContainsField(Field field)
        {
            //Iterate through fields in the collection
            for (int i = 0; i < this.Count; i++)
            {
                //Get a field and compare its range with the specified position
                Field addedField = this[i];
                if (addedField.Range.Start.ToInt() == field.Range.Start.ToInt() &&
                    addedField.Range.Start.ToInt() == field.Range.Start.ToInt())
                    return true;
            }
            return false;
        }
    }

    /// <summary>
    /// A comparer used to sort fields by their positions
    /// </summary>
    class FieldComparer : IComparer<Field>
    {
        public int Compare(Field x, Field y)
        {
            //Use the default Integer comparer to compare positions
            return Comparer<int>.Default.Compare(x.Range.Start.ToInt(), y.Range.Start.ToInt());
        }
    }
}
