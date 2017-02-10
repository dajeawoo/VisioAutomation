using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{

    public abstract class WriterBase<TValue>
    {
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        protected readonly List<SRC> SRCs;
        protected readonly List<TValue> SRC_Values;

        protected readonly List<SIDSRC> SIDSRCs;
        protected readonly List<TValue> SIDSRC_Values;

        public void Clear()
        {
            this.SRCs.Clear();
            this.SRC_Values.Clear();

            this.SIDSRCs.Clear();
            this.SIDSRC_Values.Clear();
        }

        protected void Add(SRC src, TValue value)
        {
            this.SRCs.Add(src);
            this.SRC_Values.Add(value);
        }

        protected void Add(SIDSRC sidsrc, TValue value)
        {
            this.SIDSRCs.Add(sidsrc);
            this.SIDSRC_Values.Add(value);
        }

        protected WriterBase()
        {
            this.SRCs = new List<SRC>();
            this.SRC_Values = new List<TValue>();

            this.SIDSRCs = new List<SIDSRC>();
            this.SIDSRC_Values = new List<TValue>();
        }

        protected IVisio.VisGetSetArgs ComputeGetResultFlags(ResultType rt)
        {
            var flags = this.combine_blastguards_and_testcircular_flags();

            if (rt == ResultType.ResultString)
            {
                flags |= IVisio.VisGetSetArgs.visGetStrings;
            }

            return flags;
        }

        protected IVisio.VisGetSetArgs ComputeGetFormulaFlags()
        {
            var common_flags = this.combine_blastguards_and_testcircular_flags();
            var formula_flags = (short)IVisio.VisGetSetArgs.visSetUniversalSyntax;
            var combined_flags = (short)common_flags | formula_flags;
            return (IVisio.VisGetSetArgs)combined_flags;
        }

        private IVisio.VisGetSetArgs combine_blastguards_and_testcircular_flags()
        {
            var f_bg = this.BlastGuards ? IVisio.VisGetSetArgs.visSetBlastGuards : 0;
            var f_tc = this.TestCircular ? IVisio.VisGetSetArgs.visSetTestCircular : 0;

            var flags = ((short)f_bg) | ((short)f_tc);
            return (IVisio.VisGetSetArgs)flags;
        }

        public void Commit(VisioAutomation.ShapeSheet.ShapeSheetSurface surface)
        {
            if (this.SRCCount > 0)
            {
                this.CommitSRC(surface);
            }

            if (this.SIDSRCCount > 0)
            {
                this.CommitSIDSRC(surface);
            }
        }

        protected abstract void CommitSRC(VisioAutomation.ShapeSheet.ShapeSheetSurface surface);
        protected abstract void CommitSIDSRC(VisioAutomation.ShapeSheet.ShapeSheetSurface surface);

        public int Count => this.SRC_Values.Count + this.SIDSRC_Values.Count;

        protected short[] GetSIDSRCStream()
        {
            var stream = SIDSRC.ToStream(this.SIDSRCs);
            if (stream.Length != this.SIDSRCCount*4)
            {
                throw new System.ArgumentException();
            }

            return stream;
        }

        protected short[] GetSRCStream()
        {
            var stream = SRC.ToStream(this.SRCs);
            if (stream.Length != this.SRCCount * 3)
            {
                throw new System.ArgumentException();
            }

            return stream;
        }

        protected int SIDSRCCount => this.SIDSRCs.Count;
        protected int SRCCount => this.SRCs.Count;

    }


    public enum CoordType
    {
        SIDSRC,
        SRC
    }

    public struct WriteRec<TValue>
    {
        private readonly SIDSRC SIDSRC;
        public readonly SRC SRC;
        public readonly TValue Value;
        public readonly CoordType Type;

        public WriteRec(SIDSRC sidsrc, TValue value)
        {
            this.SIDSRC = sidsrc;
            this.SRC = new SRC();
            this.Value = value;
            this.Type = CoordType.SIDSRC;
        }

        public WriteRec(SRC src, TValue value)
        {
            this.SIDSRC = new SIDSRC();
            this.SRC = src;
            this.Value = value;
            this.Type = CoordType.SRC;
        }

        public SIDSRC Sidsrc
        {
            get
            {
                if (this.Type != CoordType.SIDSRC)
                {
                    throw new System.ArgumentException();
                }
                return SIDSRC;
            }
        }

        public SRC Src
        {
            get
            {
                if (this.Type != CoordType.SRC)
                {
                    throw new System.ArgumentException();
                }
                return SRC;
            }
        }
    }

    public abstract class XWriterBase<TValue>
    {
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        protected readonly List<WriteRec<TValue>> Records;

        public void Clear()
        {
            this.Records.Clear();
        }

        protected void Add(SRC src, TValue value)
        {
            var rec = new WriteRec<TValue>(src, value);
            this.Records.Add(rec);
        }

        protected void Add(SIDSRC sidsrc, TValue value)
        {
            var rec = new WriteRec<TValue>(sidsrc, value);
            this.Records.Add(rec);
        }

        protected XWriterBase()
        {
            this.Records = new List<WriteRec<TValue>>();
        }

        protected IVisio.VisGetSetArgs ComputeGetResultFlags(ResultType rt)
        {
            var flags = this.combine_blastguards_and_testcircular_flags();

            if (rt == ResultType.ResultString)
            {
                flags |= IVisio.VisGetSetArgs.visGetStrings;
            }

            return flags;
        }

        protected IVisio.VisGetSetArgs ComputeGetFormulaFlags()
        {
            var common_flags = this.combine_blastguards_and_testcircular_flags();
            var formula_flags = (short)IVisio.VisGetSetArgs.visSetUniversalSyntax;
            var combined_flags = (short)common_flags | formula_flags;
            return (IVisio.VisGetSetArgs)combined_flags;
        }

        private IVisio.VisGetSetArgs combine_blastguards_and_testcircular_flags()
        {
            var f_bg = this.BlastGuards ? IVisio.VisGetSetArgs.visSetBlastGuards : 0;
            var f_tc = this.TestCircular ? IVisio.VisGetSetArgs.visSetTestCircular : 0;

            var flags = ((short)f_bg) | ((short)f_tc);
            return (IVisio.VisGetSetArgs)flags;
        }

        public void Commit(VisioAutomation.ShapeSheet.ShapeSheetSurface surface)
        {
            this.CommitSRC(surface);
            this.CommitSIDSRC(surface);
        }

        protected abstract void CommitSRC(VisioAutomation.ShapeSheet.ShapeSheetSurface surface);
        protected abstract void CommitSIDSRC(VisioAutomation.ShapeSheet.ShapeSheetSurface surface);

        public int Count => this.Records.Count;

        protected IEnumerable<WriteRec<TValue>> GetRecords(CoordType type)
        {
            return this.Records.Where(i => i.Type == type);
        }
    }
}
