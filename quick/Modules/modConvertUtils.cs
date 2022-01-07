using Microsoft.VisualBasic;
using static Microsoft.VisualBasic.Constants;
using static Microsoft.VisualBasic.Interaction;
using static Microsoft.VisualBasic.Strings;
using static modUtils;
using static VBExtension;



static class modConvertUtils
{
    public static string EOLComment = "";
    public static Collection mStrings = null;
    public static int nStringCnt = 0;
    public const string DeStringToken_Base1 = "STRING_";
    public const string DeStringToken_Base2 = "TOKEN_";
    public const string DeStringToken_Base = DeStringToken_Base1 + DeStringToken_Base2;
    public static string DeComment(string Str, bool Discard = false)
    {
        string _DeComment = "";
        int A = 0;
        string T = "";
        string U = "";
        string C = "";
        _DeComment = Str;
        A = InStr(Str, "'");
        if (A == 0) return _DeComment;
        while (true)
        {
            T = Left(Str, A - 1);
            U = Replace(T, "\"", "");
            if ((Len(T) - Len(U)) % 2 == 0) break;
            A = InStr(A + 1, Str, "'");
            if (A == 0) return _DeComment;
        }
        if (!Discard) EOLComment = Mid(Str, A + 1);
        _DeComment = RTrim(Left(Str, A - 1));
        return _DeComment;
    }
    public static string ReComment(string Str, bool KeepVBComments = false)
    {
        string _ReComment = "";
        string C = "";
        string Pr = "";
        Pr = (KeepVBComments ? "'" : "//");
        if (EOLComment == "") { _ReComment = Str; return _ReComment; }
        C = Pr + EOLComment;
        EOLComment = "";
        if (!IsInStr(Str, vbCrLf))
        {
            _ReComment = Str + IIf(Len(Str) == 0, "", " ") + C;
        }
        else
        {
            _ReComment = Replace(Str, vbCrLf, C + vbCrLf, 1 , 1); // Always leave on end of first line...
        }
        if (Left(LTrim(_ReComment), 2) == Pr) _ReComment = LTrim(_ReComment);
        return _ReComment;
    }
    public static void InitDeString()
    {
        mStrings = new Collection();
        nStringCnt = 0;
    }
    private static string DeStringToken(int N)
    {
        string _DeStringToken = "";
        _DeStringToken = DeStringToken_Base + Format(N, "00000");
        return _DeStringToken;
    }
    public static string DeString(string S)
    {
        string _DeString = "";
        string Q = "\"";
        string Token = "";
        int A = 0;
        int B = 0;
        int C = 0;
        string K = "";
        if (mStrings == null) InitDeString();
        // If IsInStr(S, __S1) Then Stop
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
                _DeString = DeString(S);
                return _DeString;
            }
        }
        _DeString = S;
        return _DeString;
    }
    public static string ReString(string Str, bool doConvertString = false)
    {
        string _ReString = "";
        int I = 0;
        string T = "";
        string V = "";
        for (I = 1; I <= nStringCnt; I += 1)
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
        _ReString = Str;
        return _ReString;
    }
    private static string InternalConvertString(string S)
    {
        string _InternalConvertString = "";
        S = Replace(S, "\\", "\\\\");
        S = Replace(S, "\"\"", "\\\"");
        _InternalConvertString = S;
        return _InternalConvertString;
    }

}
