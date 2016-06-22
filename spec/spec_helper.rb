require 'nokogiri'

def html_to_ooxml(html)
 source = Nokogiri::HTML(html.gsub(/>\s+</, "><"))
 xslt = Nokogiri::XSLT(File.read('spec/html_to_ooxml.xslt'))
 result = xslt.apply_to(source)
 result
end