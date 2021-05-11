using Microsoft.VisualBasic;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static modUtils;
using static VBExtension;


static class modConvertUtils
{
    // Option Explicit
    private static string EOLComment = "";
    private static Collection mStrings = null;
    private static int nStringCnt = 0;
    private const string DeStringToken_Base1 = "STRING_";
    private const string DeStringToken_Base2 = "TOKEN_";
    public const string DeStringToken_Base = DeStringToken_Base1 + DeStringToken_Base2;


    public static string DeComment(string Str, bool Discard = false)
    {
        string DeComment = "";
        int A = 0;

        string T = "";
        string U = "";

        string C = "";

        DeComment = Str;
        A = InStr(Str, "'");
        if (A == 0)
        {
            return DeComment;

        }
        while (true)
        {
            T = Left(Str, A - 1);
            U = Replace(T, "\"", "");
            if ((Len(T) - Len(U)) % 2 == 0)
            {
                break;
            }
            A = InStr(A + 1, Str, "'");
            if (A == 0)
            {
                return DeComment;

            }
        }
        if (!Discard)
        {
            EOLComment = Mid(Str, A + 1);
        }
        DeComment = RTrim(Left(Str, A - 1));
        return DeComment;
    }

    public static dynamic ReComment(string Str, bool KeepVBComments = false)
    {
        dynamic ReComment = null;
        string C = "";

        string Pr = "";

        Pr = IIf(KeepVBComments, "'", "//");
        if (EOLComment == "")
        {
            ReComment = Str;
            return ReComment;

        }
        C = Pr + EOLComment;
        EOLComment = "";
        if (!IsInStr(Str, vbCrLf))
        {
            ReComment = Str + IIf(Len(Str) == 0, "", " ") + C;
        }
        else
        {
            ReComment = Replace(Str, vbCrLf, C + vbCrLf, _, 1); // Always leave on end of first line...
        }
        if (Left(LTrim(ReComment), 2) == Pr)
        {
            ReComment = LTrim(ReComment);
        }
        return ReComment;
    }

    public static void InitDeString()
    {
        mStrings = new Collection(); ;
        nStringCnt = 0;
    }

    private static string DeStringToken(int N)
    {
        string DeStringToken = "";
        DeStringToken = DeStringToken_Base + Format(N, "00000");
        return DeStringToken;
    }

    public static string DeString(string S)
    {
        string DeString = "";
        const string Q = "\"";
        string Token = "";

        int A = 0;
        int B = 0;
        int C = 0;

        string K = "";

        if (mStrings == null)
        {
            InitDeString();
        }

        //If IsInStr(S, """ArCheck.chkShowB") Then Stop

        A = InStr(S, Q);
        C = A;
        if (A > 0)
        {
        MidQuote:;
            B = InStr(C + 1, S, Q);
            if (B > 0)
            {
                if (Mid(S, B + 1, 1) == Q)
                {
                    C = B + 1;
                    goto MidQuote;
                }
                nStringCnt = nStringCnt + 1;
                Token = DeStringToken(nStringCnt);
                K = Mid(S, A, B - A + 1);
                mStrings.Add(K, Token);
                S = Left(S, A - 1) + Token + Mid(S, B + 1);
                DeString = DeString[S];
                return DeString;

            }
        }
        DeString = S;
        return DeString;
    }

    public static string ReString(string Str, bool doConvertString = false)
    {
        string ReString = "";
        int I = 0;
        string T = "";
        string V = "";

        for (I = 1; I < nStringCnt; I++)
        {
            T = DeStringToken(I);
            V = mStrings.Item(T);
            if (V != "" && doConvertString)
            {
                if (Left(V, 1) == "\"" && Right(V, 1) == "\"")
                {
                    V = "\"" + InternalConvertString(Mid(V, 2, Len(V) - 2)) + "\"";
                }
            }
            Str = Replace(Str, T, V);
        }
        ReString = Str;
        return ReString;
    }

    private static dynamic InternalConvertString(string S)
    {
        dynamic InternalConvertString = null;
        S = Replace(S, "\\", "\\\\");
        S = Replace(S, "\"\"", "\\\"");
        InternalConvertString = S;
        return InternalConvertString;
    }
}
