import unittest

from lxml import etree

from format_figure_ref_phrases import NS, Q, format_paragraph


class FormatFigureReferencePhraseTests(unittest.TestCase):
    def test_formats_whole_phrase_across_ref_result_without_changing_field(self):
        xml = f"""<w:p xmlns:w="{NS['w']}" xmlns:xml="http://www.w3.org/XML/1998/namespace">
          <w:r><w:t>结构如</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText> REF MyFigure \\h </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>图</w:t></w:r>
          <w:r><w:t xml:space="preserve"> </w:t></w:r>
          <w:r><w:t>37</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
          <w:r><w:t>所示。其他正文</w:t></w:r>
        </w:p>"""
        paragraph = etree.fromstring(xml.encode())
        original_field = paragraph.xpath(".//w:instrText/text()", namespaces=NS)
        self.assertEqual(format_paragraph(paragraph, 28), 1)
        self.assertEqual(
            "".join(paragraph.xpath(".//w:t/text()", namespaces=NS)),
            "结构如图 37所示。其他正文",
        )
        self.assertEqual(paragraph.xpath(".//w:instrText/text()", namespaces=NS), original_field)

        content = ""
        for run in paragraph.xpath(".//w:r", namespaces=NS):
            value = "".join(run.xpath("./w:t/text()", namespaces=NS))
            size = run.find("./w:rPr/w:sz", namespaces=NS)
            if value in ("如", "图", " ", "37", "所示"):
                self.assertIsNotNone(size)
                self.assertEqual(size.get(Q("val")), "28")
            if value in ("结构", "。其他正文"):
                self.assertIsNone(size)
            content += value
        self.assertEqual(content, "结构如图 37所示。其他正文")

    def test_rejects_matched_run_with_field_marker(self):
        xml = f"""<w:p xmlns:w="{NS['w']}">
          <w:r><w:t>如图 1所示</w:t><w:fldChar w:fldCharType="end"/></w:r>
        </w:p>"""
        with self.assertRaises(ValueError):
            format_paragraph(etree.fromstring(xml.encode()), 28)


if __name__ == "__main__":
    unittest.main()
