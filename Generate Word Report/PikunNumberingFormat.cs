using System.CodeDom.Compiler;


namespace Generate_Word_Report
{
    [GeneratedCode("DomGen", "2.0")]
    public enum PikunNumberingFormat
    {
        //
        // Summary:
        //     Decimal Numbers.
        //     When the item is serialized out as xml, its value is "decimal".
        Decimal = 0,
        //
        // Summary:
        //     Uppercase Roman Numerals.
        //     When the item is serialized out as xml, its value is "upperRoman".
        UpperRoman = 1,
        //
        // Summary:
        //     Lowercase Roman Numerals.
        //     When the item is serialized out as xml, its value is "lowerRoman".
        LowerRoman = 2,
        //
        // Summary:
        //     Uppercase Latin Alphabet.
        //     When the item is serialized out as xml, its value is "upperLetter".
        UpperLetter = 3,
        //
        // Summary:
        //     Lowercase Latin Alphabet.
        //     When the item is serialized out as xml, its value is "lowerLetter".
        LowerLetter = 4,
        //
        // Summary:
        //     Ordinal.
        //     When the item is serialized out as xml, its value is "ordinal".
        Ordinal = 5,
        //
        // Summary:
        //     Cardinal Text.
        //     When the item is serialized out as xml, its value is "cardinalText".
        CardinalText = 6,
        //
        // Summary:
        //     Ordinal Text.
        //     When the item is serialized out as xml, its value is "ordinalText".
        OrdinalText = 7,
        //
        // Summary:
        //     Hexadecimal Numbering.
        //     When the item is serialized out as xml, its value is "hex".
        Hex = 8,
        //
        // Summary:
        //     Chicago Manual of Style.
        //     When the item is serialized out as xml, its value is "chicago".
        Chicago = 9,
        //
        // Summary:
        //     Ideographs.
        //     When the item is serialized out as xml, its value is "ideographDigital".
        IdeographDigital = 10,
        //
        // Summary:
        //     Japanese Counting System.
        //     When the item is serialized out as xml, its value is "japaneseCounting".
        JapaneseCounting = 11,
        //
        // Summary:
        //     AIUEO Order Hiragana.
        //     When the item is serialized out as xml, its value is "aiueo".
        Aiueo = 12,
        //
        // Summary:
        //     Iroha Ordered Katakana.
        //     When the item is serialized out as xml, its value is "iroha".
        Iroha = 13,
        //
        // Summary:
        //     Double Byte Arabic Numerals.
        //     When the item is serialized out as xml, its value is "decimalFullWidth".
        DecimalFullWidth = 14,
        //
        // Summary:
        //     Single Byte Arabic Numerals.
        //     When the item is serialized out as xml, its value is "decimalHalfWidth".
        DecimalHalfWidth = 15,
        //
        // Summary:
        //     Japanese Legal Numbering.
        //     When the item is serialized out as xml, its value is "japaneseLegal".
        JapaneseLegal = 16,
        //
        // Summary:
        //     Japanese Digital Ten Thousand Counting System.
        //     When the item is serialized out as xml, its value is "japaneseDigitalTenThousand".
        JapaneseDigitalTenThousand = 17,
        //
        // Summary:
        //     Decimal Numbers Enclosed in a Circle.
        //     When the item is serialized out as xml, its value is "decimalEnclosedCircle".
        DecimalEnclosedCircle = 18,
        //
        // Summary:
        //     Double Byte Arabic Numerals Alternate.
        //     When the item is serialized out as xml, its value is "decimalFullWidth2".
        DecimalFullWidth2 = 19,
        //
        // Summary:
        //     Full-Width AIUEO Order Hiragana.
        //     When the item is serialized out as xml, its value is "aiueoFullWidth".
        AiueoFullWidth = 20,
        //
        // Summary:
        //     Full-Width Iroha Ordered Katakana.
        //     When the item is serialized out as xml, its value is "irohaFullWidth".
        IrohaFullWidth = 21,
        //
        // Summary:
        //     Initial Zero Arabic Numerals.
        //     When the item is serialized out as xml, its value is "decimalZero".
        DecimalZero = 22,
        //
        // Summary:
        //     Bullet.
        //     When the item is serialized out as xml, its value is "bullet".
        Bullet = 23,
        //
        // Summary:
        //     Korean Ganada Numbering.
        //     When the item is serialized out as xml, its value is "ganada".
        Ganada = 24,
        //
        // Summary:
        //     Korean Chosung Numbering.
        //     When the item is serialized out as xml, its value is "chosung".
        Chosung = 25,
        //
        // Summary:
        //     Decimal Numbers Followed by a Period.
        //     When the item is serialized out as xml, its value is "decimalEnclosedFullstop".
        DecimalEnclosedFullstop = 26,
        //
        // Summary:
        //     Decimal Numbers Enclosed in Parenthesis.
        //     When the item is serialized out as xml, its value is "decimalEnclosedParen".
        DecimalEnclosedParen = 27,
        //
        // Summary:
        //     Decimal Numbers Enclosed in a Circle.
        //     When the item is serialized out as xml, its value is "decimalEnclosedCircleChinese".
        DecimalEnclosedCircleChinese = 28,
        //
        // Summary:
        //     Ideographs Enclosed in a Circle.
        //     When the item is serialized out as xml, its value is "ideographEnclosedCircle".
        IdeographEnclosedCircle = 29,
        //
        // Summary:
        //     Traditional Ideograph Format.
        //     When the item is serialized out as xml, its value is "ideographTraditional".
        IdeographTraditional = 30,
        //
        // Summary:
        //     Zodiac Ideograph Format.
        //     When the item is serialized out as xml, its value is "ideographZodiac".
        IdeographZodiac = 31,
        //
        // Summary:
        //     Traditional Zodiac Ideograph Format.
        //     When the item is serialized out as xml, its value is "ideographZodiacTraditional".
        IdeographZodiacTraditional = 32,
        //
        // Summary:
        //     Taiwanese Counting System.
        //     When the item is serialized out as xml, its value is "taiwaneseCounting".
        TaiwaneseCounting = 33,
        //
        // Summary:
        //     Traditional Legal Ideograph Format.
        //     When the item is serialized out as xml, its value is "ideographLegalTraditional".
        IdeographLegalTraditional = 34,
        //
        // Summary:
        //     Taiwanese Counting Thousand System.
        //     When the item is serialized out as xml, its value is "taiwaneseCountingThousand".
        TaiwaneseCountingThousand = 35,
        //
        // Summary:
        //     Taiwanese Digital Counting System.
        //     When the item is serialized out as xml, its value is "taiwaneseDigital".
        TaiwaneseDigital = 36,
        //
        // Summary:
        //     Chinese Counting System.
        //     When the item is serialized out as xml, its value is "chineseCounting".
        ChineseCounting = 37,
        //
        // Summary:
        //     Chinese Legal Simplified Format.
        //     When the item is serialized out as xml, its value is "chineseLegalSimplified".
        ChineseLegalSimplified = 38,
        //
        // Summary:
        //     Chinese Counting Thousand System.
        //     When the item is serialized out as xml, its value is "chineseCountingThousand".
        ChineseCountingThousand = 39,
        //
        // Summary:
        //     Korean Digital Counting System.
        //     When the item is serialized out as xml, its value is "koreanDigital".
        KoreanDigital = 40,
        //
        // Summary:
        //     Korean Counting System.
        //     When the item is serialized out as xml, its value is "koreanCounting".
        KoreanCounting = 41,
        //
        // Summary:
        //     Korean Legal Numbering.
        //     When the item is serialized out as xml, its value is "koreanLegal".
        KoreanLegal = 42,
        //
        // Summary:
        //     Korean Digital Counting System Alternate.
        //     When the item is serialized out as xml, its value is "koreanDigital2".
        KoreanDigital2 = 43,
        //
        // Summary:
        //     Vietnamese Numerals.
        //     When the item is serialized out as xml, its value is "vietnameseCounting".
        VietnameseCounting = 44,
        //
        // Summary:
        //     Lowercase Russian Alphabet.
        //     When the item is serialized out as xml, its value is "russianLower".
        RussianLower = 45,
        //
        // Summary:
        //     Uppercase Russian Alphabet.
        //     When the item is serialized out as xml, its value is "russianUpper".
        RussianUpper = 46,
        //
        // Summary:
        //     No Numbering.
        //     When the item is serialized out as xml, its value is "none".
        None = 47,
        //
        // Summary:
        //     Number With Dashes.
        //     When the item is serialized out as xml, its value is "numberInDash".
        NumberInDash = 48,
        //
        // Summary:
        //     Hebrew Numerals.
        //     When the item is serialized out as xml, its value is "hebrew1".
        Hebrew1 = 49,
        //
        // Summary:
        //     Hebrew Alphabet.
        //     When the item is serialized out as xml, its value is "hebrew2".
        Hebrew2 = 50,
        //
        // Summary:
        //     Arabic Alphabet.
        //     When the item is serialized out as xml, its value is "arabicAlpha".
        ArabicAlpha = 51,
        //
        // Summary:
        //     Arabic Abjad Numerals.
        //     When the item is serialized out as xml, its value is "arabicAbjad".
        ArabicAbjad = 52,
        //
        // Summary:
        //     Hindi Vowels.
        //     When the item is serialized out as xml, its value is "hindiVowels".
        HindiVowels = 53,
        //
        // Summary:
        //     Hindi Consonants.
        //     When the item is serialized out as xml, its value is "hindiConsonants".
        HindiConsonants = 54,
        //
        // Summary:
        //     Hindi Numbers.
        //     When the item is serialized out as xml, its value is "hindiNumbers".
        HindiNumbers = 55,
        //
        // Summary:
        //     Hindi Counting System.
        //     When the item is serialized out as xml, its value is "hindiCounting".
        HindiCounting = 56,
        //
        // Summary:
        //     Thai Letters.
        //     When the item is serialized out as xml, its value is "thaiLetters".
        ThaiLetters = 57,
        //
        // Summary:
        //     Thai Numerals.
        //     When the item is serialized out as xml, its value is "thaiNumbers".
        ThaiNumbers = 58,
        //
        // Summary:
        //     Thai Counting System.
        //     When the item is serialized out as xml, its value is "thaiCounting".
        ThaiCounting = 59,
        //
        // Summary:
        //     bahtText.
        //     When the item is serialized out as xml, its value is "bahtText".
        BahtText = 60,
        //
        // Summary:
        //     dollarText.
        //     When the item is serialized out as xml, its value is "dollarText".
        DollarText = 61,
        //
        // Summary:
        //     custom.
        //     When the item is serialized out as xml, its value is "custom".
        Custom = 62
    }
}

