﻿using DocumentFormat.OpenXml;
using System;
using System.Drawing;
using System.IO;

namespace Pikunword
{
    public static class Pikun
    {
        public static string horizontalAlignmentLeft = "left";
        public static string horizontalAlignmentCenter = "center";
        public static string horizontalAlignmentRight = "right";

        public static string paragraphUnderline = "U";
        public static string paragraphItalic = "I";
        public static string paragraphBold = "B";
        public static string paragraphNormal = "N";

        public static string wrapSquare = "WrapSquare";
        public static string wrapTopBottom = "WrapTopBottom";
        public static string wrapNone = "WrapNone";
        public static string wrapThrough = "wrapThrough";
        public static string wrapTight = "WrapTight";

        public static int numberingIdDecimal = 1;
        public static int numberingIdBullet = 2;

        public static string justificationLeft = "left";
        public static string justificationCenter = "center";
        public static string justificationRight = "right";

        public static string highlightColorBlack = "Black";
        public static string highlightColorBlue = "Blue";
        public static string highlightColorCyan = "Cyan";
        public static string highlightColorGreen = "Green";
        public static string highlightColorMagenta = "Magenta";
        public static string highlightColorRed = "Red";
        public static string highlightColorYellow = "Yellow";
        public static string highlightColorWhite = "White";
        public static string highlightColorDarkBlue = "DarkBlue";
        public static string highlightColorDarkCyan = "DarkCyan";
        public static string highlightColorDarkGreen = "DarkGreen";
        public static string highlightColorDarkMagenta = "DarkMagenta";
        public static string highlightColorDarkRed = "DarkRed";
        public static string highlightColorDarkYellow = "DarkYellow";
        public static string highlightColorDarkGray = "DarkGray";
        public static string highlightColorLightGray = "LightGray";
        public static string highlightColorNone = "None";

        public static string numberFormatValuesBullet = "Bullet";
        public static string numberFormatValuesDecimal = "Decimal";
        public static string numberFormatValuesDecimalABC = "DecimalABC";

        public static string tableCellVerticalAlignmentTop = "Top";
        public static string tableCellVerticalAlignmentCenter = "Center";
        public static string tableCellVerticalAlignmentBottom = "Bottom";


        #region color
        public static string Black_W3C = "000000";
        public static string Night = "0C090A";
        public static string Charcoal = "34282C";
        public static string Oil = "3B3131";
        public static string Dark_Gray = "3A3B3C";
        public static string Light_Black = "454545";
        public static string Black_Cat = "413839";
        public static string Iridium = "3D3C3A";
        public static string Black_Eel = "463E3F";
        public static string Black_Cow = "4C4646";
        public static string Gray_Wolf = "504A4B";
        public static string Vampire_Gray = "565051";
        public static string Iron_Gray = "52595D";
        public static string Gray_Dolphin = "5C5858";
        public static string Carbon_Gray = "625D5D";
        public static string Ash_Gray = "666362";
        public static string Cloudy_Gray = "6D6968";
        public static string DimGray_or_DimGrey_W3C = "696969";
        public static string Smokey_Gray = "726E6D";
        public static string Alien_Gray = "736F6E";
        public static string Sonic_Silver = "757575";
        public static string Platinum_Gray = "797979";
        public static string Granite = "837E7C";
        public static string Gray_or_Grey_W3C = "808080";
        public static string Battleship_Gray = "848482";
        public static string DarkGray_or_DarkGrey_W3C = "A9A9A9";
        public static string Gray_Cloud = "B6B6B4";
        public static string Silver_W3C = "C0C0C0";
        public static string Pale_Silver = "C9C0BB";
        public static string Gray_Goose = "D1D0CE";
        public static string Platinum_Silver = "CECECE";
        public static string LightGray_or_LightGrey_W3C = "D3D3D3";
        public static string Gainsboro_W3C = "DCDCDC";
        public static string Platinum = "E5E4E2";
        public static string Metallic_Silver = "BCC6CC";
        public static string Blue_Gray = "98AFC7";
        public static string Roman_Silver = "838996";
        public static string LightSlateGray_or_LightSlateGrey_W3C = "778899";
        public static string SlateGray_or_SlateGrey_W3C = "708090";
        public static string Rat_Gray = "6D7B8D";
        public static string Slate_Granite_Gray = "657383";
        public static string Jet_Gray = "616D7E";
        public static string Mist_Blue = "646D7E";
        public static string Marble_Blue = "566D7E";
        public static string Slate_Blue_Grey = "737CA1";
        public static string Light_Purple_Blue = "728FCE";
        public static string Azure_Blue = "4863A0";
        public static string Blue_Jay = "2B547E";
        public static string Charcoal_Blue = "36454F";
        public static string Dark_Blue_Grey = "29465B";
        public static string Dark_Slate = "2B3856";
        public static string Deep_Sea_Blue = "123456";
        public static string Night_Blue = "151B54";
        public static string MidnightBlue_W3C = "191970";
        public static string Navy_W3C = "000080";
        public static string Denim_Dark_Blue = "151B8D";
        public static string DarkBlue_W3C = "00008B";
        public static string Lapis_Blue = "15317E";
        public static string New_Midnight_Blue = "0000A0";
        public static string Earth_Blue = "0000A5";
        public static string Cobalt_Blue = "0020C2";
        public static string MediumBlue_W3C = "0000CD";
        public static string Blueberry_Blue = "0041C2";
        public static string Canary_Blue = "2916F5";
        public static string Blue_W3C = "0000FF";
        public static string Bright_Blue = "0909FF";
        public static string Blue_Orchid = "1F45FC";
        public static string Sapphire_Blue = "2554C7";
        public static string Blue_Eyes = "1569C7";
        public static string Bright_Navy_Blue = "1974D2";
        public static string Balloon_Blue = "2B60DE";
        public static string RoyalBlue_W3C = "4169E1";
        public static string Ocean_Blue = "2B65EC";
        public static string Blue_Ribbon = "306EFF";
        public static string Blue_Dress = "157DEC";
        public static string Neon_Blue = "1589FF";
        public static string DodgerBlue_W3C = "1E90FF";
        public static string Glacial_Blue_Ice = "368BC1";
        public static string SteelBlue_W3C = "4682B4";
        public static string Silk_Blue = "488AC7";
        public static string Windows_Blue = "357EC7";
        public static string Blue_Ivy = "3090C7";
        public static string Blue_Koi = "659EC7";
        public static string Columbia_Blue = "87AFC7";
        public static string Baby_Blue = "95B9C7";
        public static string CornflowerBlue_W3C = "6495ED";
        public static string Sky_Blue_Dress = "6698FF";
        public static string Iceberg = "56A5EC";
        public static string Butterfly_Blue = "38ACEC";
        public static string DeepSkyBlue_W3C = "00BFFF";
        public static string Midday_Blue = "3BB9FF";
        public static string Crystal_Blue = "5CB3FF";
        public static string Denim_Blue = "79BAEC";
        public static string Day_Sky_Blue = "82CAFF";
        public static string LightSkyBlue_W3C = "87CEFA";
        public static string SkyBlue_W3C = "87CEEB";
        public static string Jeans_Blue = "A0CFEC";
        public static string Blue_Angel = "B7CEEC";
        public static string Pastel_Blue = "B4CFEC";
        public static string Light_Day_Blue = "ADDFFF";
        public static string Sea_Blue = "C2DFFF";
        public static string Heavenly_Blue = "C6DEFF";
        public static string Robin_Egg_Blue = "BDEDFF";
        public static string PowderBlue_W3C = "B0E0E6";
        public static string Coral_Blue = "AFDCEC";
        public static string LightBlue_W3C = "ADD8E6";
        public static string LightSteelBlue_W3C = "B0CFDE";
        public static string Gulf_Blue = "C9DFEC";
        public static string Pastel_Light_Blue = "D5D6EA";
        public static string Lavender_Blue = "E3E4FA";
        public static string Lavender_W3C = "E6E6FA";
        public static string Water = "EBF4FA";
        public static string AliceBlue_W3C = "F0F8FF";
        public static string GhostWhite_W3C = "F8F8FF";
        public static string Azure_W3C = "F0FFFF";
        public static string LightCyan_W3C = "E0FFFF";
        public static string Light_Slate = "CCFFFF";
        public static string Electric_Blue = "9AFEFF";
        public static string Tron_Blue = "7DFDFE";
        public static string Blue_Zircon = "57FEFF";
        public static string Aqua_or_Cyan_W3C = "00FFFF";
        public static string Bright_Cyan = "0AFFFF";
        public static string Celeste = "50EBEC";
        public static string Blue_Diamond = "4EE2EC";
        public static string Bright_Turquoise = "16E2F5";
        public static string Blue_Lagoon = "8EEBEC";
        public static string PaleTurquoise_W3C = "AFEEEE";
        public static string Pale_Blue_Lily = "CFECEC";
        public static string Tiffany_Blue = "81D8D0";
        public static string Blue_Hosta = "77BFC7";
        public static string Cyan_Opaque = "92C7C7";
        public static string Northern_Lights_Blue = "78C7C7";
        public static string Blue_Green = "7BCCB5";
        public static string MediumAquaMarine_W3C = "66CDAA";
        public static string Magic_Mint = "AAF0D1";
        public static string Aquamarine_W3C = "7FFFD4";
        public static string Light_Aquamarine = "93FFE8";
        public static string Turquoise_W3C = "40E0D0";
        public static string MediumTurquoise_W3C = "48D1CC";
        public static string Deep_Turquoise = "48CCCD";
        public static string Jellyfish = "46C7C7";
        public static string Blue_Turquoise = "43C6DB";
        public static string DarkTurquoise_W3C = "00CED1";
        public static string Macaw_Blue_Green = "43BFC7";
        public static string LightSeaGreen_W3C = "20B2AA";
        public static string Seafoam_Green = "3EA99F";
        public static string CadetBlue_W3C = "5F9EA0";
        public static string Deep_Sea = "3B9C9C";
        public static string DarkCyan_W3C = "008B8B";
        public static string Teal_W3C = "008080";
        public static string Medium_Teal = "045F5F";
        public static string Deep_Teal = "033E3E";
        public static string DarkSlateGray_or_DarkSlateGrey_W3C = "25383C";
        public static string Gunmetal = "2C3539";
        public static string Blue_Moss_Green = "3C565B";
        public static string Beetle_Green = "4C787E";
        public static string Grayish_Turquoise = "5E7D7E";
        public static string Greenish_Blue = "307D7E";
        public static string Aquamarine_Stone = "348781";
        public static string Sea_Turtle_Green = "438D80";
        public static string Dull_Sea_Green = "4E8975";
        public static string Deep_Sea_Green = "306754";
        public static string SeaGreen_W3C = "2E8B57";
        public static string Dark_Mint = "31906E";
        public static string Jade = "00A36C";
        public static string Earth_Green = "34A56F";
        public static string Emerald = "50C878";
        public static string Mint = "3EB489";
        public static string MediumSeaGreen_W3C = "3CB371";
        public static string Camouflage_Green = "78866B";
        public static string Sage_Green = "848B79";
        public static string Hazel_Green = "617C58";
        public static string Venom_Green = "728C00";
        public static string OliveDrab_W3C = "6B8E23";
        public static string Olive_W3C = "808000";
        public static string DarkOliveGreen_W3C = "556B2F";
        public static string Army_Green = "4B5320";
        public static string Fern_Green = "667C26";
        public static string Fall_Forest_Green = "4E9258";
        public static string Pine_Green = "387C44";
        public static string Medium_Forest_Green = "347235";
        public static string Jungle_Green = "347C2C";
        public static string ForestGreen_W3C = "228B22";
        public static string Green_W3C = "008000";
        public static string DarkGreen_W3C = "006400";
        public static string Deep_Emerald_Green = "046307";
        public static string Dark_Forest_Green = "254117";
        public static string Seaweed_Green = "437C17";
        public static string Shamrock_Green = "347C17";
        public static string Green_Onion = "6AA121";
        public static string Green_Pepper = "4AA02C";
        public static string Dark_Lime_Green = "41A317";
        public static string Parrot_Green = "12AD2B";
        public static string Clover_Green = "3EA055";
        public static string Dinosaur_Green = "73A16C";
        public static string Green_Snake = "6CBB3C";
        public static string Alien_Green = "6CC417";
        public static string Green_Apple = "4CC417";
        public static string LimeGreen_W3C = "32CD32";
        public static string Pea_Green = "52D017";
        public static string Kelly_Green = "4CC552";
        public static string Zombie_Green = "54C571";
        public static string Frog_Green = "99C68E";
        public static string DarkSeaGreen_W3C = "8FBC8F";
        public static string Green_Peas = "89C35C";
        public static string Dollar_Bill_Green = "85BB65";
        public static string Iguana_Green = "9CB071";
        public static string Acid_Green = "B0BF1A";
        public static string Avocado_Green = "B2C248";
        public static string Pistachio_Green = "9DC209";
        public static string Salad_Green = "A1C935";
        public static string YellowGreen_W3C = "9ACD32";
        public static string Pastel_Green = "77DD77";
        public static string Hummingbird_Green = "7FE817";
        public static string Nebula_Green = "59E817";
        public static string Stoplight_Go_Green = "57E964";
        public static string Neon_Green = "16F529";
        public static string Jade_Green = "5EFB6E";
        public static string Lime_Mint_Green = "36F57F";
        public static string SpringGreen_W3C = "00FF7F";
        public static string MediumSpringGreen_W3C = "00FA9A";
        public static string Emerald_Green = "5FFB17";
        public static string Lime_W3C = "00FF00";
        public static string LawnGreen_W3C = "7CFC00";
        public static string Bright_Green = "66FF00";
        public static string Chartreuse_W3C = "7FFF00";
        public static string Yellow_Lawn_Green = "87F717";
        public static string Aloe_Vera_Green = "98F516";
        public static string Dull_Green_Yellow = "B1FB17";
        public static string GreenYellow_W3C = "ADFF2F";
        public static string Chameleon_Green = "BDF516";
        public static string Neon_Yellow_Green = "DAEE01";
        public static string Yellow_Green_Grosbeak = "E2F516";
        public static string Tea_Green = "CCFB5D";
        public static string Slime_Green = "BCE954";
        public static string Algae_Green = "64E986";
        public static string LightGreen_W3C = "90EE90";
        public static string Dragon_Green = "6AFB92";
        public static string PaleGreen_W3C = "98FB98";
        public static string Mint_Green = "98FF98";
        public static string Green_Thumb = "B5EAAA";
        public static string Organic_Brown = "E3F9A6";
        public static string Light_Jade = "C3FDB8";
        public static string Light_Rose_Green = "DBF9DB";
        public static string HoneyDew_W3C = "F0FFF0";
        public static string MintCream_W3C = "F5FFFA";
        public static string LemonChiffon_W3C = "FFFACD";
        public static string Parchment = "FFFFC2";
        public static string Cream = "FFFFCC";
        public static string LightGoldenRodYellow_W3C = "FAFAD2";
        public static string LightYellow_W3C = "FFFFE0";
        public static string Beige_W3C = "F5F5DC";
        public static string Cornsilk_W3C = "FFF8DC";
        public static string Blonde = "FBF6D9";
        public static string Champagne = "F7E7CE";
        public static string AntiqueWhite_W3C = "FAEBD7";
        public static string PapayaWhip_W3C = "FFEFD5";
        public static string BlanchedAlmond_W3C = "FFEBCD";
        public static string Bisque_W3C = "FFE4C4";
        public static string Wheat_W3C = "F5DEB3";
        public static string Moccasin_W3C = "FFE4B5";
        public static string Peach = "FFE5B4";
        public static string Light_Orange = "FED8B1";
        public static string PeachPuff_W3C = "FFDAB9";
        public static string NavajoWhite_W3C = "FFDEAD";
        public static string Golden_Blonde = "FBE7A1";
        public static string Golden_Silk = "F3E3C3";
        public static string Dark_Blonde = "F0E2B6";
        public static string Light_Gold = "F1E5AC";
        public static string Vanilla = "F3E5AB";
        public static string Tan_Brown = "ECE5B6";
        public static string PaleGoldenRod_W3C = "EEE8AA";
        public static string Khaki_W3C = "F0E68C";
        public static string Cardboard_Brown = "EDDA74";
        public static string Harvest_Gold = "EDE275";
        public static string Sun_Yellow = "FFE87C";
        public static string Corn_Yellow = "FFF380";
        public static string Pastel_Yellow = "FAF884";
        public static string Neon_Yellow = "FFFF33";
        public static string Yellow_W3C = "FFFF00";
        public static string Canary_Yellow = "FFEF00";
        public static string Banana_Yellow = "F5E216";
        public static string Mustard_Yellow = "FFDB58";
        public static string Golden_Yellow = "FFDF00";
        public static string Bold_Yellow = "F9DB24";
        public static string Rubber_Ducky_Yellow = "FFD801";
        public static string Gold_W3C = "FFD700";
        public static string Bright_Gold = "FDD017";
        public static string Golden_Brown = "EAC117";
        public static string Deep_Yellow = "F6BE00";
        public static string Macaroni_and_Cheese = "F2BB66";
        public static string Saffron = "FBB917";
        public static string Beer = "FBB117";
        public static string Yellow_Orange_or_Orange_Yellow = "FFAE42";
        public static string Cantaloupe = "FFA62F";
        public static string Orange_W3C = "FFA500";
        public static string Brown_Sand = "EE9A4D";
        public static string SandyBrown_W3C = "F4A460";
        public static string Brown_Sugar = "E2A76F";
        public static string Camel_Brown = "C19A6B";
        public static string Deer_Brown = "E6BF83";
        public static string BurlyWood_W3C = "DEB887";
        public static string Tan_W3C = "D2B48C";
        public static string Light_French_Beige = "C8AD7F";
        public static string Sand = "C2B280";
        public static string Sage = "BCB88A";
        public static string Fall_Leaf_Brown = "C8B560";
        public static string Ginger_Brown = "C9BE62";
        public static string DarkKhaki_W3C = "BDB76B";
        public static string Olive_Green = "BAB86C";
        public static string Brass = "B5A642";
        public static string Cookie_Brown = "C7A317";
        public static string Metallic_Gold = "D4AF37";
        public static string Bee_Yellow = "E9AB17";
        public static string School_Bus_Yellow = "E8A317";
        public static string GoldenRod_W3C = "DAA520";
        public static string Orange_Gold = "D4A017";
        public static string Caramel = "C68E17";
        public static string DarkGoldenRod_W3C = "B8860B";
        public static string Cinnamon = "C58917";
        public static string Peru_W3C = "CD853F";
        public static string Bronze = "CD7F32";
        public static string Tiger_Orange = "C88141";
        public static string Copper = "B87333";
        public static string Wood = "966F33";
        public static string Oak_Brown = "806517";
        public static string Antique_Bronze = "665D1E";
        public static string Hazel = "8E7618";
        public static string Dark_Yellow = "8B8000";
        public static string Dark_Moccasin = "827839";
        public static string Bullet_Shell = "AF9B60";
        public static string Army_Brown = "827B60";
        public static string Sandstone = "786D5F";
        public static string Taupe = "483C32";
        public static string Mocha = "493D26";
        public static string Milk_Chocolate = "513B1C";
        public static string Gray_Brown = "3D3635";
        public static string Dark_Coffee = "3B2F2F";
        public static string Old_Burgundy = "43302E";
        public static string Western_Charcoal = "49413F";
        public static string Bakers_Brown = "5C3317";
        public static string Dark_Brown = "654321";
        public static string Sepia_Brown = "704214";
        public static string Coffee = "6F4E37";
        public static string Brown_Bear = "835C3B";
        public static string Red_Dirt = "7F5217";
        public static string Sepia = "7F462C";
        public static string Sienna_W3C = "A0522D";
        public static string SaddleBrown_W3C = "8B4513";
        public static string Dark_Sienna = "8A4117";
        public static string Sangria = "7E3817";
        public static string Blood_Red = "7E3517";
        public static string Chestnut = "954535";
        public static string Chestnut_Red = "C34A2C";
        public static string Mahogany = "C04000";
        public static string Red_Fox = "C35817";
        public static string Dark_Bisque = "B86500";
        public static string Light_Brown = "B5651D";
        public static string Rust = "C36241";
        public static string Copper_Red = "CB6D51";
        public static string Orange_Salmon = "C47451";
        public static string Chocolate_W3C = "D2691E";
        public static string Sedona = "CC6600";
        public static string Papaya_Orange = "E56717";
        public static string Halloween_Orange = "E66C2C";
        public static string Neon_Orange = "FF6700";
        public static string Bright_Orange = "FF5F1F";
        public static string Pumpkin_Orange = "F87217";
        public static string Carrot_Orange = "F88017";
        public static string DarkOrange_W3C = "FF8C00";
        public static string Construction_Cone_Orange = "F87431";
        public static string Indian_Saffron = "FF7722";
        public static string Sunrise_Orange = "E67451";
        public static string Mango_Orange = "FF8040";
        public static string Coral_W3C = "FF7F50";
        public static string Basket_Ball_Orange = "F88158";
        public static string Light_Salmon_Rose = "F9966B";
        public static string LightSalmon_W3C = "FFA07A";
        public static string DarkSalmon_W3C = "E9967A";
        public static string Tangerine = "E78A61";
        public static string Light_Copper = "DA8A67";
        public static string Salmon_W3C = "FA8072";
        public static string LightCoral_W3C = "F08080";
        public static string Pastel_Red = "F67280";
        public static string Pink_Coral = "E77471";
        public static string Bean_Red = "F75D59";
        public static string Valentine_Red = "E55451";
        public static string IndianRed_W3C = "CD5C5C";
        public static string Tomato_W3C = "FF6347";
        public static string Shocking_Orange = "E55B3C";
        public static string OrangeRed_W3C = "FF4500";
        public static string Red_W3C = "FF0000";
        public static string Neon_Red = "FD1C03";
        public static string Scarlet = "FF2400";
        public static string Ruby_Red = "F62217";
        public static string Ferrari_Red = "F70D1A";
        public static string Fire_Engine_Red = "F62817";
        public static string Lava_Red = "E42217";
        public static string Love_Red = "E41B17";
        public static string Grapefruit = "DC381F";
        public static string Cherry_Red = "C24641";
        public static string Chilli_Pepper = "C11B17";
        public static string FireBrick_W3C = "B22222";
        public static string Tomato_Sauce_Red = "B21807";
        public static string Brown_W3C = "A52A2A";
        public static string Carbon_Red = "A70D2A";
        public static string Cranberry = "9F000F";
        public static string Saffron_Red = "931314";
        public static string Red_Wine_or_Wine_Red = "990012";
        public static string DarkRed_W3C = "8B0000";
        public static string Maroon_W3C = "800000";
        public static string Burgundy = "8C001A";
        public static string Deep_Red = "800517";
        public static string Red_Blood = "660000";
        public static string Blood_Night = "551606";
        public static string Black_Bean = "3D0C02";
        public static string Chocolate_Brown = "3F000F";
        public static string Midnight = "2B1B17";
        public static string Purple_Lily = "550A35";
        public static string Purple_Maroon = "810541";
        public static string Plum_Pie = "7D0541";
        public static string Plum_Velvet = "7D0552";
        public static string Dark_Raspberry = "872657";
        public static string Velvet_Maroon = "7E354D";
        public static string Rosy_Finch = "7F4E52";
        public static string Dull_Purple = "7F525D";
        public static string Puce = "7F5A58";
        public static string Rose_Dust = "997070";
        public static string Rosy_Pink = "B38481";
        public static string RosyBrown_W3C = "BC8F8F";
        public static string Khaki_Rose = "C5908E";
        public static string Pink_Brown = "C48189";
        public static string Lipstick_Pink = "C48793";
        public static string Rose = "E8ADAA";
        public static string Silver_Pink = "C4AEAD";
        public static string Rose_Gold = "ECC5C0";
        public static string Deep_Peach = "FFCBA4";
        public static string Pastel_Orange = "F8B88B";
        public static string Desert_Sand = "EDC9AF";
        public static string Unbleached_Silk = "FFDDCA";
        public static string Pig_Pink = "FDD7E4";
        public static string Blush = "FFE6E8";
        public static string MistyRose_W3C = "FFE4E1";
        public static string Pink_Bubble_Gum = "FFDFDD";
        public static string Light_Red = "FFCCCB";
        public static string Light_Rose = "FBCFCD";
        public static string Deep_Rose = "FBBBB9";
        public static string Pink_W3C = "FFC0CB";
        public static string LightPink_W3C = "FFB6C1";
        public static string Donut_Pink = "FAAFBE";
        public static string Baby_Pink = "FAAFBA";
        public static string Flamingo_Pink = "F9A7B0";
        public static string Pastel_Pink = "FEA3AA";
        public static string Pink_Rose = "E7A1B0";
        public static string Pink_Daisy = "E799A3";
        public static string Cadillac_Pink = "E38AAE";
        public static string Carnation_Pink = "F778A1";
        public static string Blush_Red = "E56E94";
        public static string PaleVioletRed_W3C = "DB7093";
        public static string Purple_Pink = "D16587";
        public static string Tulip_Pink = "C25A7C";
        public static string Bashful_Pink = "C25283";
        public static string Dark_Pink = "E75480";
        public static string Dark_Hot_Pink = "F660AB";
        public static string HotPink_W3C = "FF69B4";
        public static string Watermelon_Pink = "FC6C85";
        public static string Violet_Red = "F6358A";
        public static string Hot_Deep_Pink = "F52887";
        public static string DeepPink_W3C = "FF1493";
        public static string Neon_Pink = "F535AA";
        public static string Neon_Hot_Pink = "FD349C";
        public static string Pink_Cupcake = "E45E9D";
        public static string Dimorphotheca_Magenta = "E3319D";
        public static string Pink_Lemonade = "E4287C";
        public static string Raspberry = "E30B5D";
        public static string Crimson_W3C = "DC143C";
        public static string Bright_Maroon = "C32148";
        public static string Rose_Red = "C21E56";
        public static string Rogue_Pink = "C12869";
        public static string Burnt_Pink = "C12267";
        public static string Pink_Violet = "CA226B";
        public static string MediumVioletRed_W3C = "C71585";
        public static string Dark_Carnation_Pink = "C12283";
        public static string Raspberry_Purple = "B3446C";
        public static string Pink_Plum = "B93B8F";
        public static string Orchid_W3C = "DA70D6";
        public static string Deep_Mauve = "DF73D4";
        public static string Violet_W3C = "EE82EE";
        public static string Bright_Neon_Pink = "F433FF";
        public static string Fuchsia_or_Magenta_W3C = "FF00FF";
        public static string Crimson_Purple = "E238EC";
        public static string Heliotrope_Purple = "D462FF";
        public static string Tyrian_Purple = "C45AEC";
        public static string MediumOrchid_W3C = "BA55D3";
        public static string Purple_Flower = "A74AC7";
        public static string Orchid_Purple = "B048B5";
        public static string Pastel_Violet = "D291BC";
        public static string Mauve_Taupe = "915F6D";
        public static string Viola_Purple = "7E587E";
        public static string Eggplant = "614051";
        public static string Plum_Purple = "583759";
        public static string Grape = "5E5A80";
        public static string Purple_Navy = "4E5180";
        public static string SlateBlue_W3C = "6A5ACD";
        public static string Blue_Lotus = "6960EC";
        public static string Light_Slate_Blue = "736AFF";
        public static string MediumSlateBlue_W3C = "7B68EE";
        public static string Periwinkle_Purple = "7575CF";
        public static string Purple_Amethyst = "6C2DC7";
        public static string Bright_Purple = "6A0DAD";
        public static string Deep_Periwinkle = "5453A6";
        public static string DarkSlateBlue_W3C = "483D8B";
        public static string Purple_Haze = "4E387E";
        public static string Purple_Iris = "571B7E";
        public static string Dark_Purple = "4B0150";
        public static string Deep_Purple = "36013F";
        public static string Purple_Monster = "461B7E";
        public static string Indigo_W3C = "4B0082";
        public static string Blue_Whale = "342D7E";
        public static string RebeccaPurple_W3C = "663399";
        public static string Purple_Jam = "6A287E";
        public static string DarkMagenta_W3C = "8B008B";
        public static string Purple_W3C = "800080";
        public static string French_Lilac = "86608E";
        public static string DarkOrchid_W3C = "9932CC";
        public static string DarkViolet_W3C = "9400D3";
        public static string Purple_Violet = "8D38C9";
        public static string Jasmine_Purple = "A23BEC";
        public static string Purple_Daffodil = "B041FF";
        public static string Clemantis_Violet = "842DCE";
        public static string BlueViolet_W3C = "8A2BE2";
        public static string Purple_Sage_Bush = "7A5DC7";
        public static string Lovely_Purple = "7F38EC";
        public static string Neon_Purple = "9D00FF";
        public static string Purple_Plum = "8E35EF";
        public static string Aztech_Purple = "893BFF";
        public static string Lavender_Purple = "967BB6";
        public static string MediumPurple_W3C = "9370DB";
        public static string Light_Purple = "8467D7";
        public static string Crocus_Purple = "9172EC";
        public static string Purple_Mimosa = "9E7BFF";
        public static string Periwinkle = "CCCCFF";
        public static string Pale_Lilac = "DCD0FF";
        public static string Mauve = "E0B0FF";
        public static string Bright_Lilac = "D891EF";
        public static string Rich_Lilac = "B666D2";
        public static string Purple_Dragon = "C38EC7";
        public static string Lilac = "C8A2C8";
        public static string Plum_W3C = "DDA0DD";
        public static string Blush_Pink = "E6A9EC";
        public static string Pastel_Purple = "F2A2E8";
        public static string Blossom_Pink = "F9B7FF";
        public static string Wisteria_Purple = "C6AEC7";
        public static string Purple_Thistle = "D2B9D3";
        public static string Thistle_W3C = "D8BFD8";
        public static string Periwinkle_Pink = "E9CFEC";
        public static string Cotton_Candy = "FCDFFF";
        public static string Lavender_Pinocchio = "EBDDE2";
        public static string Ash_White = "E9E4D4";
        public static string White_Chocolate = "EDE6D6";
        public static string Soft_Ivory = "FAF0DD";
        public static string Off_White = "F8F0E3";
        public static string LavenderBlush_W3C = "FFF0F5";
        public static string Pearl = "FDEEF4";
        public static string Egg_Shell = "FFF9E3";
        public static string OldLace_W3C = "FDF5E6";
        public static string Linen_W3C = "FAF0E6";
        public static string SeaShell_W3C = "FFF5EE";
        public static string Rice = "FAF5EF";
        public static string FloralWhite_W3C = "FFFAF0";
        public static string Ivory_W3C = "FFFFF0";
        public static string Light_White = "FFFFF7";
        public static string WhiteSmoke_W3C = "F5F5F5";
        public static string Cotton = "FBFBF9";
        public static string Snow_W3C = "FFFAFA";
        public static string Milk_White = "FEFCFF";
        public static string White_W3C = "FFFFFF";
        #endregion

        public static bool IsNullableType(Type type)
        {
            return type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);
        }
        public static bool IsNullOrEmpty(this object data)
        {
            return string.IsNullOrEmpty(Convert.ToString(data));
        }
        public static int AsInt(this object data, int? defaultValue = null)
        {
            if (IsNullOrEmpty(data))
                return defaultValue != null ? Convert.ToInt32(defaultValue) : 0;

            return Convert.ToInt32(data);
        }
        public static Bitmap Base64StringToBitmap(this string base64String)
        {
            Bitmap bmpReturn = null;

            byte[] byteBuffer = Convert.FromBase64String(base64String);
            MemoryStream memoryStream = new MemoryStream(byteBuffer);

            memoryStream.Position = 0;

            bmpReturn = (Bitmap)Bitmap.FromStream(memoryStream);

            memoryStream.Close();
            memoryStream = null;
            byteBuffer = null;


            return bmpReturn;
        }
        public static string BitmapToBase64String(this string path)
        {
            try
            {
                using (Image image = Image.FromFile(path))
                {
                    using (MemoryStream m = new MemoryStream())
                    {
                        image.Save(m, image.RawFormat);
                        byte[] imageBytes = m.ToArray();

                        // Convert byte[] to Base64 String
                        string base64String = Convert.ToBase64String(imageBytes);
                        return base64String;
                    }
                }
            }
            catch (Exception e)
            {
                return "";
            }
            
        }
        public static UInt32Value AsUInt32Value(this int i)
        {
           return i != 0 ? UInt32Value.FromUInt32((UInt32)(i)): UInt32Value.FromUInt32((UInt32)(0));
        }
    }
}
