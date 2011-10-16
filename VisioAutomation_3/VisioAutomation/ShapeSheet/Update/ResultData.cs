﻿using System.Collections;
using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Update
{
    public class ResultData<TStream> : IEnumerable<ResultItem<TStream>> where TStream : struct
    {
        private readonly List<ResultItem<TStream>> items;

        public ResultData()
        {
            this.items = new List<ResultItem<TStream>>();
        }

        public ResultData(int capacity)
        {
            this.items = new List<ResultItem<TStream>>(capacity);
        }

        public int Count
        {
            get { return this.items.Count; }
        }

        public void Set(TStream streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            var rec = new ResultItem<TStream>(streamitem, value, unitcode);
            this.items.Add(rec);
        }

        public double[] GetResultsArray()
        {
            return ShapeSheetHelper.MapCollectionToArray(this.items, r => r.Result);
        }

        public IVisio.VisUnitCodes[] GetUnitCodesArray()
        {
            return ShapeSheetHelper.MapCollectionToArray(this.items, r => r.UnitCode);
        }

        public IEnumerator<ResultItem<TStream>> GetEnumerator()
        {
            foreach (var i in this.items)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     // Explicit implementation
        {                                           // keeps it hidden.
            return GetEnumerator();
        }
    }
}