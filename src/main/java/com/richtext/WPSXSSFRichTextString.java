package com.richtext;

import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;

import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.reflect.Field;

/**
 * <p>用于解决POI导出的Excel中富文本加粗不生效的问题，</p>
 * <p>此问题是由于较低版本的WPS不能正常识别 &lt;b val="true"/&gt;</p>
 */
public class WPSXSSFRichTextString extends XSSFRichTextString {
    MethodHandles.Lookup lookup = MethodHandles.lookup();
    private static Field stFeild;

    static {
        try {
            stFeild = XSSFRichTextString.class.getDeclaredField("st");
            stFeild.setAccessible(true);
        } catch (NoSuchFieldException e) {
            throw new RuntimeException(e);
        }
    }


    @Override
    public void append(String text, XSSFFont font) {
        try {

            MethodHandle stGetter = lookup.unreflectGetter(stFeild);
            CTRst st = (CTRst) stGetter.invoke(this);
            if (st.sizeOfRArray() == 0 && st.isSetT()) {
                // convert <t>string</t> into a text run: <r><t>string</t></r>
                CTRElt lt = st.addNewR();
                lt.setT(st.getT());
                preserveSpaces(lt.xgetT());
                st.unsetT();
            }
            CTRElt lt = st.addNewR();
            lt.setT(text);
            preserveSpaces(lt.xgetT());

            if (font != null) {
                CTRPrElt pr = lt.addNewRPr();
                setRunAttributes(font.getCTFont(), pr);
            }
        } catch (Throwable e) {
            throw new RuntimeException(e);
        }
    }

    private void setRunAttributes(CTFont ctFont, CTRPrElt pr) {
        if (ctFont.sizeOfBArray() > 0 && ctFont.getBArray(0).getVal()) pr.addNewB();
        if (ctFont.sizeOfUArray() > 0) pr.addNewU().setVal(ctFont.getUArray(0).getVal());
        if (ctFont.sizeOfIArray() > 0 && ctFont.getIArray(0).getVal()) pr.addNewI();
        if (ctFont.sizeOfColorArray() > 0) {
            CTColor c1 = ctFont.getColorArray(0);
            CTColor c2 = pr.addNewColor();
            if (c1.isSetAuto()) c2.setAuto(c1.getAuto());
            if (c1.isSetIndexed()) c2.setIndexed(c1.getIndexed());
            if (c1.isSetRgb()) c2.setRgb(c1.getRgb());
            if (c1.isSetTheme()) c2.setTheme(c1.getTheme());
            if (c1.isSetTint()) c2.setTint(c1.getTint());
        }
        if (ctFont.sizeOfSzArray() > 0) pr.addNewSz().setVal(ctFont.getSzArray(0).getVal());
        if (ctFont.sizeOfNameArray() > 0) pr.addNewRFont().setVal(ctFont.getNameArray(0).getVal());
        if (ctFont.sizeOfFamilyArray() > 0) pr.addNewFamily().setVal(ctFont.getFamilyArray(0).getVal());
        if (ctFont.sizeOfSchemeArray() > 0) pr.addNewScheme().setVal(ctFont.getSchemeArray(0).getVal());
        if (ctFont.sizeOfCharsetArray() > 0) pr.addNewCharset().setVal(ctFont.getCharsetArray(0).getVal());

        if (ctFont.sizeOfCondenseArray() > 0 && ctFont.getCondenseArray(0).getVal()) pr.addNewCondense();
        if (ctFont.sizeOfExtendArray() > 0 && ctFont.getExtendArray(0).getVal()) pr.addNewExtend();
        if (ctFont.sizeOfVertAlignArray() > 0) pr.addNewVertAlign().setVal(ctFont.getVertAlignArray(0).getVal());
        if (ctFont.sizeOfOutlineArray() > 0 && ctFont.getOutlineArray(0).getVal()) pr.addNewOutline();
        if (ctFont.sizeOfShadowArray() > 0 && ctFont.getShadowArray(0).getVal()) pr.addNewShadow();
        if (ctFont.sizeOfStrikeArray() > 0 && ctFont.getStrikeArray(0).getVal()) pr.addNewStrike();
    }
}
