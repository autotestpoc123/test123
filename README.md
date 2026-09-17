 using System.Xml; using System.IO;
  var xml = "<Root><Text2>7G754</Text2></Root>";
  using var r = XmlReader.Create(new StringReader(xml), new XmlReaderSettings{IgnoreWhitespace=true});
  r.MoveToContent();  r.Read();                 // 进入 <Text2>
  while (!(r.NodeType==XmlNodeType.EndElement && r.LocalName=="Root")) {
      if (r.NodeType==XmlNodeType.Element) { r.Read(); continue; }   // ← 原版逻辑
      if (r.LocalName=="Text2") Console.WriteLine(r.ReadElementContentAsString());
      else r.Skip();
  }
