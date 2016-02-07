﻿using System;
using System.IO;
using System.Xml.Serialization;
using ImageProcessor.Imaging.Filters;
using PowerPointLabs.PictureSlidesLab.Service.Effect;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    [Serializable]
    public class StyleOptions : WPF.Observable.Model
    {
        public StyleOptions()
        {
            Init();
        }

        # region UI related prop
        private bool _isUseTextFormat;

        public bool IsUseTextFormat
        {
            get
            {
                return _isUseTextFormat;
            }
            set
            {
                _isUseTextFormat = value;
                OnPropertyChanged("IsUseTextFormat");
            }
        }

        private string _fontFamily;

        public string FontFamily
        {
            get { return _fontFamily; }
            set
            {
                _fontFamily = value;
                OnPropertyChanged("FontFamily");
            }
        }

        private int _fontSizeIncrease;

        public int FontSizeIncrease
        {
            get { return _fontSizeIncrease; }
            set
            {
                _fontSizeIncrease = value;
                OnPropertyChanged("FontSizeIncrease");
            }
        }

        private string _fontColor;

        public string FontColor
        {
            get { return _fontColor; }
            set
            {
                _fontColor = value;
                OnPropertyChanged("FontColor");
            }
        }

        private int _textBoxPosition;

        public int TextBoxPosition
        {
            get { return _textBoxPosition; }
            set
            {
                _textBoxPosition = value;
                OnPropertyChanged("TextBoxPosition");
            }
        }

        private int _textBoxAlignment;

        public int TextBoxAlignment
        {
            get { return _textBoxAlignment; }
            set
            {
                _textBoxAlignment = value;
                OnPropertyChanged("TextBoxAlignment");
            }
        }

        private int _imageOffset;

        public int ImageOffset
        {
            get { return _imageOffset; }
            set
            {
                _imageOffset = value;
                OnPropertyChanged("ImageOffset");
            }
        }

        // ******************************************************
        // for overlay style
        // ******************************************************

        private bool _isUseOverlayStyle;

        public bool IsUseOverlayStyle
        {
            get { return _isUseOverlayStyle; }
            set
            {
                _isUseOverlayStyle = value;
                OnPropertyChanged("IsUseOverlayStyle");
            }
        }

        private string _overlayColor;

        public string OverlayColor
        {
            get { return _overlayColor; }
            set
            {
                _overlayColor = value;
                OnPropertyChanged("OverlayColor");
            }
        }

        private int _transparency;

        // for background's overlay
        public int Transparency
        {
            get { return _transparency; }
            set
            {
                _transparency = value;
                OnPropertyChanged("Transparency");
            }
        }

        // ******************************************************
        // for textbox style
        // ******************************************************

        private bool _isUseTextBoxStyle;

        public bool IsUseTextBoxStyle
        {
            get { return _isUseTextBoxStyle; }
            set
            {
                _isUseTextBoxStyle = value;
                OnPropertyChanged("IsUseTextBoxStyle");
            }
        }

        private string _textBoxColor;

        public string TextBoxColor
        {
            get { return _textBoxColor; }
            set
            {
                _textBoxColor = value;
                OnPropertyChanged("TextBoxColor");
            }
        }

        private int _textBoxTransparency;

        public int TextBoxTransparency
        {
            get { return _textBoxTransparency; }
            set
            {
                _textBoxTransparency = value;
                OnPropertyChanged("TextBoxTransparency");
            }
        }

        // ******************************************************
        // for banner style
        // ******************************************************

        private bool _isUseBannerStyle;

        public bool IsUseBannerStyle
        {
            get { return _isUseBannerStyle; }
            set
            {
                _isUseBannerStyle = value;
                OnPropertyChanged("IsUseBannerStyle");
            }
        }

        private int _bannerShape;

        public int BannerShape
        {
            get { return _bannerShape; }
            set
            {
                _bannerShape = value;
                OnPropertyChanged("BannerShape");
            }
        }

        private int _bannerDirection;

        public int BannerDirection
        {
            get { return _bannerDirection; }
            set
            {
                _bannerDirection = value;
                OnPropertyChanged("BannerDirection");
            }
        }

        private string _bannerColor;

        public string BannerColor
        {
            get { return _bannerColor; }
            set
            {
                _bannerColor = value;
                OnPropertyChanged("BannerColor");
            }
        }

        private int _bannerTransparency;

        public int BannerTransparency
        {
            get { return _bannerTransparency; }
            set
            {
                _bannerTransparency = value;
                OnPropertyChanged("BannerTransparency");
            }
        }

        // ******************************************************
        // for special effect style
        // ******************************************************

        private bool _isUseSpecialEffectStyle;

        public bool IsUseSpecialEffectStyle
        {
            get { return _isUseSpecialEffectStyle; }
            set
            {
                _isUseSpecialEffectStyle = value;
                OnPropertyChanged("IsUseSpecialEffectStyle");
            }
        }

        private int _specialEffect;

        public int SpecialEffect
        {
            get { return _specialEffect; }
            set
            {
                _specialEffect = value;
                OnPropertyChanged("SpecialEffect");
            }
        }

        // ******************************************************
        // for blur style
        // ******************************************************

        private bool _isUseBlurStyle;

        public bool IsUseBlurStyle
        {
            get { return _isUseBlurStyle; }
            set
            {
                _isUseBlurStyle = value;
                OnPropertyChanged("IsUseBlurStyle");
            }
        }

        private int _blurDegree;

        public int BlurDegree
        {
            get { return _blurDegree; }
            set
            {
                _blurDegree = value;
                OnPropertyChanged("BlurDegree");
            }
        }

        // ******************************************************
        // for outline style
        // ******************************************************

        private bool _isUseOutlineStyle;

        public bool IsUseOutlineStyle
        {
            get { return _isUseOutlineStyle; }
            set
            {
                _isUseOutlineStyle = value;
                OnPropertyChanged("IsUseOutlineStyle");
            }
        }

        // ******************************************************
        // for frame style
        // ******************************************************

        private bool _isUseFrameStyle;

        public bool IsUseFrameStyle
        {
            get { return _isUseFrameStyle; }
            set
            {
                _isUseFrameStyle = value;
                OnPropertyChanged("IsUseFrameStyle");
            }
        }

        private string _frameColor;

        public string FrameColor
        {
            get { return _frameColor; }
            set
            {
                _frameColor = value;
                OnPropertyChanged("FrameColor");
            }
        }

        private int _frameTransparency;

        public int FrameTransparency
        {
            get { return _frameTransparency; }
            set
            {
                _frameTransparency = value;
                OnPropertyChanged("FrameTransparency");
            }
        }

        // ******************************************************
        // for circle style
        // ******************************************************

        private bool _isUseCircleStyle;

        public bool IsUseCircleStyle
        {
            get { return _isUseCircleStyle; }
            set
            {
                _isUseCircleStyle = value;
                OnPropertyChanged("IsUseCircleStyle");
            }
        }

        private string _circleColor;

        public string CircleColor
        {
            get { return _circleColor; }
            set
            {
                _circleColor = value;
                OnPropertyChanged("CircleColor");
            }
        }

        private int _circleTransparency;

        public int CircleTransparency
        {
            get { return _circleTransparency; }
            set
            {
                _circleTransparency = value;
                OnPropertyChanged("CircleTransparency");
            }
        }

        // ******************************************************
        // for triangle style
        // ******************************************************

        private bool _isUseTriangleStyle;

        public bool IsUseTriangleStyle
        {
            get { return _isUseTriangleStyle; }
            set
            {
                _isUseTriangleStyle = value;
                OnPropertyChanged("IsUseTriangleStyle");
            }
        }

        private string _triangleColor;

        public string TriangleColor
        {
            get { return _triangleColor; }
            set
            {
                _triangleColor = value;
                OnPropertyChanged("TriangleColor");
            }
        }

        private int _triangleTransparency;

        public int TriangleTransparency
        {
            get { return _triangleTransparency; }
            set
            {
                _triangleTransparency = value;
                OnPropertyChanged("TriangleTransparency");
            }
        }

        // *****************************************
        // Text Glow
        // *****************************************

        private bool _isUseTextGlow;

        public bool IsUseTextGlow
        {
            get { return _isUseTextGlow; }
            set
            {
                _isUseTextGlow = value;
                OnPropertyChanged("IsUseTextGlow");
            }
        }

        private string _textGlowColor;

        public string TextGlowColor
        {
            get { return _textGlowColor; }
            set
            {
                _textGlowColor = value;
                OnPropertyChanged("TextGlowColor");
            }
        }

        // other

        private bool _isInsertReference;

        public bool IsInsertReference
        {
            get { return _isInsertReference; }
            set
            {
                _isInsertReference = value;
                OnPropertyChanged("IsInsertReference");
            }
        }

        // used as tooltip for Variation Stage
        public string OptionName { get; set; }

        // used for Reload Styles
        public string StyleName { get; set; }

        // used for Reload Styles
        public int VariantIndex { get; set; }

        # endregion

        # region Logic
        public void Init()
        {
            ImageOffset = 0;

            IsUseTextFormat = true;
            FontFamily = "Calibri";
            FontSizeIncrease = 0;
            FontColor = "#FFFFFF";
            TextBoxPosition = 5;
            TextBoxAlignment = 0;

            IsUseOverlayStyle = false;
            OverlayColor = "#000000";
            Transparency = 100;

            IsUseBannerStyle = false;
            BannerColor = "#000000";
            BannerTransparency = 25;
            BannerShape = 0;
            BannerDirection = 0;

            IsUseTextBoxStyle = false;
            TextBoxColor = "#000000";
            TextBoxTransparency = 25;

            IsUseSpecialEffectStyle = false;
            SpecialEffect = -1;

            IsUseBlurStyle = false;
            BlurDegree = 0;

            IsUseOutlineStyle = false;

            IsUseFrameStyle = false;
            FrameColor = "#FFFFFF";
            FrameTransparency = 30;

            IsUseCircleStyle = false;
            CircleColor = "#FFFFFF";
            CircleTransparency = 0;

            IsUseTriangleStyle = false;
            TriangleColor = "#000000";
            TriangleTransparency = 0;

            IsUseTextGlow = false;
            TextGlowColor = "#123456";

            IsInsertReference = false;
            OptionName = "Default";
            StyleName = "";
            VariantIndex = 0;
        }

        public string GetFontFamily()
        {
            return FontFamily;
        }

        public Position GetTextBoxPosition()
        {
            switch (TextBoxPosition)
            {
                case 0:
                    return Position.Original;
                case 1:
                    return Position.TopLeft;
                case 2:
                    return Position.Top;
                case 3:
                    return Position.TopRight;
                case 4:
                    return Position.Left;
                case 5:
                    return Position.Centre;
                case 6:
                    return Position.Right;
                case 7:
                    return Position.BottomLeft;
                case 8:
                    return Position.Bottom;
                // case 9:
                default:
                    return Position.BottomRight;
            }
        }

        public Alignment GetTextBoxAlignment()
        {
            switch (TextBoxAlignment)
            {
                case 0:
                    return Alignment.Auto;
                case 1:
                    return Alignment.Left;
                case 2:
                    return Alignment.Centre;
                // case 3:
                default:
                    return Alignment.Right;
            }
        }

        public BannerShape GetBannerShape()
        {
            switch (BannerShape)
            {
                case 0:
                    return Service.Effect.BannerShape.Rectangle;
                case 1:
                    return Service.Effect.BannerShape.Circle;
                case 2:
                    return Service.Effect.BannerShape.RectangleOutline;
                default:
                    return Service.Effect.BannerShape.CircleOutline;
            }
        }

        public BannerDirection GetBannerDirection()
        {
            switch (BannerDirection)
            {
                case 0:
                    return Service.Effect.BannerDirection.Auto;
                case 1:
                    return Service.Effect.BannerDirection.Horizontal;
                // case 2:
                default:
                    return Service.Effect.BannerDirection.Vertical;
            }
        }

        public IMatrixFilter GetSpecialEffect()
        {
            switch (SpecialEffect)
            {
                case 0:
                    return MatrixFilters.GreyScale;
                case 1:
                    return MatrixFilters.BlackWhite;
                case 2:
                    return MatrixFilters.Comic;
                case 3:
                    return MatrixFilters.Gotham;
                case 4:
                    return MatrixFilters.HiSatch;
                case 5:
                    return MatrixFilters.Invert;
                case 6:
                    return MatrixFilters.Lomograph;
                case 7:
                    return MatrixFilters.LoSatch;
                case 8:
                    return MatrixFilters.Polaroid;
                // case 9:
                default:
                    return MatrixFilters.Sepia;
            }
        }

        # endregion

        # region IO serialization
        /// Taken from http://stackoverflow.com/a/14663848

        /// <summary>
        /// Saves to an xml file
        /// </summary>
        /// <param name="filename">File path of the new xml file</param>
        public void Save(string filename)
        {
            try
            {
                using (var writer = new StreamWriter(filename))
                {
                    var serializer = new XmlSerializer(GetType());
                    serializer.Serialize(writer, this);
                    writer.Flush();
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.Log("Failed to save Picture Slides Lab Style Options: " + e.StackTrace, "Error");
            }
        }

        /// <summary>
        /// Load an object from an xml file
        /// </summary>
        /// <param name="filename">Xml file name</param>
        /// <returns>The object created from the xml file</returns>
        public static StyleOptions Load(string filename)
        {
            try
            {
                using (var stream = File.OpenRead(filename))
                {
                    var serializer = new XmlSerializer(typeof(StyleOptions));
                    var opt = serializer.Deserialize(stream) as StyleOptions;
                    return opt ?? new StyleOptions();
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.Log("Failed to load Picture Slides Lab Style Options: " + e.StackTrace, "Error");
                return new StyleOptions();
            }
        }

        # endregion
    }
}
