using System;
using System.Collections.Generic;

namespace ExtendOffice.Office.Entity
{
    public enum RecordType : uint
    {
        DocumentContainer = 0x03E8,
        DocumentAtom = 0x03E9,
        EndDocumentAtom = 0x03EA,
        SlideContainer = 0x03EE,
        SlideAtom = 0x03EF,
        Notes = 0x03F0,
        NotesAtom = 0x03F1,
        Environment = 0x03F2,
        SlidePersistAtom = 0x03F3,
        MainMaster = 0x03F8,
        SlideShowSlideInfoAtom = 0x03F9,
        SlideViewInfo = 0x03FA,
        GuideAtom = 0x03FB,
        ViewInfoAtom = 0x03FD,
        SlideViewInfoAtom = 0x03FE,
        VbaInfo = 0x03FF,
        VbaInfoAtom = 0x0400,
        SlideShowDocInfoAtom = 0x0401,
        Summary = 0x0402,
        DocRoutingSlipAtom = 0x0406,
        OutlineViewInfo = 0x0407,
        SorterViewInfo = 0x0408,
        ExternalObjectList = 0x0409,
        ExternalObjectListAtom = 0x040A,
        DrawingGroup = 0x040B,
        Drawing = 0x040C,
        GridSpacing10Atom = 0x040D,
        RoundTripTheme12Atom = 0x040E,
        RoundTripColorMapping12Atom = 0x040F,
        NamedShows = 0x0410,
        NamedShow = 0x0411,
        NamedShowSlidesAtom = 0x0412,
        NotesTextViewInfo9 = 0x0413,
        NormalViewSetInfo9 = 0x0414,
        NormalViewSetInfo9Atom = 0x0415,
        RoundTripOriginalMainMasterId12Atom = 0x041C,
        RoundTripCompositeMasterId12Atom = 0x041D,
        RoundTripContentMasterInfo12Atom = 0x041E,
        RoundTripShapeId12Atom = 0x041F,
        RoundTripHFPlaceholder12Atom = 0x0420,
        RoundTripContentMasterId12Atom = 0x0422,
        RoundTripOArtTextStyles12Atom = 0x0423,
        RoundTripHeaderFooterDefaults12Atom = 0x0424,
        RoundTripDocFlags12Atom = 0x0425,
        RoundTripShapeCheckSumForCL12Atom = 0x0426,
        RoundTripNotesMasterTextStyles12Atom = 0x0427,
        RoundTripCustomTableStyles12Atom = 0x0428,
        List = 0x07D0,
        FontCollection = 0x07D5,
        FontCollection10 = 0x07D6,
        BookmarkCollection = 0x07E3,
        SoundCollection = 0x07E4,
        SoundCollectionAtom = 0x07E5,
        Sound = 0x07E6,
        SoundDataBlob = 0x07E7,
        BookmarkSeedAtom = 0x07E9,
        ColorSchemeAtom = 0x07F0,
        BlipCollection9 = 0x07F8,
        BlipEntity9Atom = 0x07F9,
        ExternalObjectRefAtom = 0x0BC1,
        PlaceholderAtom = 0x0BC3,
        ShapeAtom = 0x0BDB,
        ShapeFlags10Atom = 0x0BDC,
        RoundTripNewPlaceholderId12Atom = 0x0BDD,
        OutlineTextRefAtom = 0x0F9E,
        TextHeaderAtom = 0x0F9F,
        TextCharsAtom = 0x0FA0,
        StyleTextPropAtom = 0x0FA1,
        MasterTextPropAtom = 0x0FA2,
        TextMasterStyleAtom = 0x0FA3,
        TextCharFormatExceptionAtom = 0x0FA4,
        TextParagraphFormatExceptionAtom = 0x0FA5,
        TextRulerAtom = 0x0FA6,
        TextBookmarkAtom = 0x0FA7,
        TextBytesAtom = 0x0FA8,
        TextSpecialInfoDefaultAtom = 0x0FA9,
        TextSpecialInfoAtom = 0x0FAA,
        DefaultRulerAtom = 0x0FAB,
        StyleTextProp9Atom = 0x0FAC,
        TextMasterStyle9Atom = 0x0FAD,
        OutlineTextProps9 = 0x0FAE,
        OutlineTextPropsHeader9Atom = 0x0FAF,
        TextDefaults9Atom = 0x0FB0,
        StyleTextProp10Atom = 0x0FB1,
        TextMasterStyle10Atom = 0x0FB2,
        OutlineTextProps10 = 0x0FB3,
        TextDefaults10Atom = 0x0FB4,
        OutlineTextProps11 = 0x0FB5,
        StyleTextProp11Atom = 0x0FB6,
        FontEntityAtom = 0x0FB7,
        FontEmbedDataBlob = 0x0FB8,
        CString = 0x0FBA,
        MetaFile = 0x0FC1,
        ExternalOleObjectAtom = 0x0FC3,
        Kinsoku = 0x0FC8,
        Handout = 0x0FC9,
        ExternalOleEmbed = 0x0FCC,
        ExternalOleEmbedAtom = 0x0FCD,
        ExternalOleLink = 0x0FCE,
        BookmarkEntityAtom = 0x0FD0,
        ExternalOleLinkAtom = 0x0FD1,
        KinsokuAtom = 0x0FD2,
        ExternalHyperlinkAtom = 0x0FD3,
        ExternalHyperlink = 0x0FD7,
        SlideNumberMetaCharAtom = 0x0FD8,
        HeadersFooters = 0x0FD9,
        HeadersFootersAtom = 0x0FDA,
        TextInteractiveInfoAtom = 0x0FDF,
        ExternalHyperlink9 = 0x0FE4,
        RecolorInfoAtom = 0x0FE7,
        ExternalOleControl = 0x0FEE,
        SlideListWithText = 0x0FF0,
        AnimationInfoAtom = 0x0FF1,
        InteractiveInfo = 0x0FF2,
        InteractiveInfoAtom = 0x0FF3,
        UserEditAtom = 0x0FF5,
        CurrentUserAtom = 0x0FF6,
        DateTimeMetaCharAtom = 0x0FF7,
        GenericDateMetaCharAtom = 0x0FF8,
        HeaderMetaCharAtom = 0x0FF9,
        FooterMetaCharAtom = 0x0FFA,
        ExternalOleControlAtom = 0x0FFB,
        ExternalMediaAtom = 0x1004,
        ExternalVideo = 0x1005,
        ExternalAviMovie = 0x1006,
        ExternalMciMovie = 0x1007,
        ExternalMidiAudio = 0x100D,
        ExternalCdAudio = 0x100E,
        ExternalWavAudioEmbedded = 0x100F,
        ExternalWavAudioLink = 0x1010,
        ExternalOleObjectStg = 0x1011,
        ExternalCdAudioAtom = 0x1012,
        ExternalWavAudioEmbeddedAtom = 0x1013,
        AnimationInfo = 0x1014,
        RtfDateTimeMetaCharAtom = 0x1015,
        ExternalHyperlinkFlagsAtom = 0x1018,
        ProgTags = 0x1388,
        ProgStringTag = 0x1389,
        ProgBinaryTag = 0x138A,
        BinaryTagDataBlob = 0x138B,
        PrintOptionsAtom = 0x1770,
        PersistDirectoryAtom = 0x1772,
        PresentationAdvisorFlags9Atom = 0x177A,
        HtmlDocInfo9Atom = 0x177B,
        HtmlPublishInfoAtom = 0x177C,
        HtmlPublishInfo9 = 0x177D,
        BroadcastDocInfo9 = 0x177E,
        BroadcastDocInfo9Atom = 0x177F,
        EnvelopeFlags9Atom = 0x1784,
        EnvelopeData9Atom = 0x1785,
        VisualShapeAtom = 0x2AFB,
        HashCodeAtom = 0x2B00,
        VisualPageAtom = 0x2B01,
        BuildList = 0x2B02,
        BuildAtom = 0x2B03,
        ChartBuild = 0x2B04,
        ChartBuildAtom = 0x2B05,
        DiagramBuild = 0x2B06,
        DiagramBuildAtom = 0x2B07,
        ParaBuild = 0x2B08,
        ParaBuildAtom = 0x2B09,
        LevelInfoAtom = 0x2B0A,
        RoundTripAnimationAtom12Atom = 0x2B0B,
        RoundTripAnimationHashAtom12Atom = 0x2B0D,
        Comment10 = 0x2EE0,
        Comment10Atom = 0x2EE1,
        CommentIndex10 = 0x2EE4,
        CommentIndex10Atom = 0x2EE5,
        LinkedShape10Atom = 0x2EE6,
        LinkedSlide10Atom = 0x2EE7,
        SlideFlags10Atom = 0x2EEA,
        SlideTime10Atom = 0x2EEB,
        DiffTree10 = 0x2EEC,
        Diff10 = 0x2EED,
        Diff10Atom = 0x2EEE,
        SlideListTableSize10Atom = 0x2EEF,
        SlideListEntry10Atom = 0x2EF0,
        SlideListTable10 = 0x2EF1,
        CryptSession10Container = 0x2F14,
        FontEmbedFlags10Atom = 0x32C8,
        FilterPrivacyFlags10Atom = 0x36B0,
        DocToolbarStates10Atom = 0x36B1,
        PhotoAlbumInfo10Atom = 0x36B2,
        SmartTagStore11Container = 0x36B3,
        RoundTripSlideSyncInfo12 = 0x3714,
        RoundTripSlideSyncInfoAtom12 = 0x3715,
        TimeConditionContainer = 0xF125,
        TimeNode = 0xF127,
        TimeCondition = 0xF128,
        TimeModifier = 0xF129,
        TimeBehaviorContainer = 0xF12A,
        TimeAnimateBehaviorContainer = 0xF12B,
        TimeColorBehaviorContainer = 0xF12C,
        TimeEffectBehaviorContainer = 0xF12D,
        TimeMotionBehaviorContainer = 0xF12E,
        TimeRotationBehaviorContainer = 0xF12F,
        TimeScaleBehaviorContainer = 0xF130,
        TimeSetBehaviorContainer = 0xF131,
        TimeCommandBehaviorContainer = 0xF132,
        TimeBehavior = 0xF133,
        TimeAnimateBehavior = 0xF134,
        TimeColorBehavior = 0xF135,
        TimeEffectBehavior = 0xF136,
        TimeMotionBehavior = 0xF137,
        TimeRotationBehavior = 0xF138,
        TimeScaleBehavior = 0xF139,
        TimeSetBehavior = 0xF13A,
        TimeCommandBehavior = 0xF13B,
        TimeClientVisualElement = 0xF13C,
        TimePropertyList = 0xF13D,
        TimeVariantList = 0xF13E,
        TimeAnimationValueList = 0xF13F,
        TimeIterateData = 0xF140,
        TimeSequenceData = 0xF141,
        TimeVariant = 0xF142,
        TimeAnimationValue = 0xF143,
        TimeExtTimeNodeContainer = 0xF144,
        TimeSubEffectContainer = 0xF145,
        OfficeArtSplitMenuColorContainer = 0xF11E,
        OfficeArtDggContainer = 0xF000,
        OfficeArtBStoreContainer = 0xF001,
        OfficeArtDgContainer = 0xF002,
        OfficeArtSpgrContainer = 0xF003,
        OfficeArtSpContainer = 0xF004,
        OfficeArtSolverContainer = 0xF005,
        OfficeArtFDGGBlock = 0xF006,
        OfficeArtInlineSpContainer = 0xF007,
        OfficeArtFDG = 0xF008,
        OfficeArtFSPGR = 0xF009,
        OfficeArtFSP = 0xF00A,
        OfficeArtFOPT = 0xF00B,
        OfficeArtClientTextbox = 0xF00D,
        OfficeArtChildAnchor = 0xF00F,
        OfficeArtClientAnchor = 0xF010,
        OfficeArtClientData = 0xF011,
        OfficeArtTertiaryFOPT = 0xF122
    }
    public class Fopte
    {
        private UInt16 _pID;
        private Byte _fBid;
        private Byte _fComplex;
        private Int32 _op;
        public UInt16 PID
        {
            get
            {
                return _pID;
            }
        }
        public Byte FBid
        {
            get
            {
                return _fBid;
            }
        }
        public Byte FComplex
        {
            get
            {
                return _fComplex;
            }
        }
        public Int32 Op
        {
            get
            {
                return _op;
            }
        }

        public Fopte(UInt16 PID, Int32 Op)
        {
            this._pID = (UInt16)(PID & 0x7FFF);
            this._fBid = (Byte)(PID >> 14);
            this._fComplex = (Byte)(PID >> 15);
            this._op = Op;
        }
    }
    public class SlideAtom
    {
        UInt32 _geom;
        UInt64 _rgPlaceholderTypes;
        UInt32 _masterIdRef;
        UInt32 _notesIdRef;
        UInt32 _slideFlags;
        UInt32 _unused;//忽略

        public UInt32 Geom
        {
            get
            {
                return _geom;
            }

            set
            {
                _geom = value;
            }
        }
        public UInt64 RgPlaceholderTypes
        {
            get
            {
                return _rgPlaceholderTypes;
            }

            set
            {
                _rgPlaceholderTypes = value;
            }
        }
        public UInt32 MasterIdRef
        {
            get
            {
                return _masterIdRef;
            }

            set
            {
                _masterIdRef = value;
            }
        }
        public UInt32 NotesIdRef
        {
            get
            {
                return _notesIdRef;
            }

            set
            {
                _notesIdRef = value;
            }
        }
        public UInt32 SlideFlags
        {
            get
            {
                return _slideFlags;
            }

            set
            {
                _slideFlags = value;
            }
        }
        public UInt32 Unused
        {
            get
            {
                return _unused;
            }

            set
            {
                _unused = value;
            }
        }

        public SlideAtom(UInt32 geom, UInt64 rgPlaceholderTypes, UInt32 masterIdRef, UInt32 notesIdRef, UInt32 slideFlags, UInt32 unused)
        {
            this._geom = geom;
            this._rgPlaceholderTypes = rgPlaceholderTypes;
            this._masterIdRef = masterIdRef;
            this._notesIdRef = notesIdRef;
            this._slideFlags = slideFlags;
            this._unused = unused;
        }
    }
    public class PowerPointRecord
    {
        #region 字段
        private UInt16 _recVer;
        private UInt16 _recInstance;
        private RecordType _recType;
        private UInt32 _recLen;
        private Int64 _offset;

        private Int32 _deepth;
        private PowerPointRecord _parent;
        private List<PowerPointRecord> _children;
        #endregion

        #region 属性
        /// <summary>
        /// 获取RecordVersion
        /// </summary>
        public UInt16 RecordVersion
        {
            get { return this._recVer; }
        }

        /// <summary>
        /// 获取RecordInstance
        /// </summary>
        public UInt16 RecordInstance
        {
            get { return this._recInstance; }
        }

        /// <summary>
        /// 获取Record类型
        /// </summary>
        public RecordType RecordType
        {
            get { return this._recType; }
        }

        /// <summary>
        /// 获取Record内容大小
        /// </summary>
        public UInt32 RecordLength
        {
            get { return this._recLen; }
        }

        /// <summary>
        /// 获取Record相对PowerPoint Document偏移
        /// </summary>
        public Int64 Offset
        {
            get { return this._offset; }
        }

        /// <summary>
        /// 获取Record深度
        /// </summary>
        public Int32 Deepth
        {
            get { return this._deepth; }
        }

        /// <summary>
        /// 获取Record的父节点
        /// </summary>
        public PowerPointRecord Parent
        {
            get { return this._parent; }
        }

        /// <summary>
        /// 获取Record的子节点
        /// </summary>
        public List<PowerPointRecord> Children
        {
            get { return this._children; }
        }
        #endregion

        #region 构造函数
        /// <summary>
        /// 初始化新的Record
        /// </summary>
        /// <param name="parent">父节点</param>
        /// <param name="version">RecordVersion和Instance</param>
        /// <param name="type">Record类型</param>
        /// <param name="length">Record内容大小</param>
        /// <param name="offset">Record相对PowerPoint Document偏移</param>
        public PowerPointRecord(PowerPointRecord parent, UInt16 version, UInt16 type, UInt32 length, Int64 offset)
        {
            this._recVer = (UInt16)(version & 0xF);
            this._recInstance = (UInt16)(version >> 4);
            this._recType = (RecordType)type;
            this._recLen = length;
            this._offset = offset;
            this._deepth = (parent == null ? 0 : parent._deepth + 1);
            this._parent = parent;

            if (_recVer == 0xF)
            {
                this._children = new List<PowerPointRecord>();
            }
        }
        #endregion

        #region 方法
        public void AddChild(PowerPointRecord entry)
        {
            if (this._children == null)
            {
                this._children = new List<PowerPointRecord>();
            }

            this._children.Add(entry);
        }
        #endregion
    }
}