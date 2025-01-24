require 'nokogiri'
require 'htmltoooxml'

include Htmltoooxml::XSLTHelper


# def html_to_ooxml(html)
#   source = Nokogiri::HTML(html.gsub(/>\s+</, '><'))
#   result = Htmltoooxml::Document.new().transform_doc_xml(source, false)
#   result.gsub!(/\s*<!--(.*?)-->\s*/m, '')
#   result = remove_declaration(result)
#   puts result
#   result
# end

def remove_whitespace(ooxml)
  ooxml.gsub(/<a:t([^>]*)>[\s\t\n\r]*/,'<a:t\1>')
end

def remove_declaration(ooxml)
  ooxml.sub(/<\?xml (.*?)>/, '').gsub(/\s*xmlns:(\w+)="(.*?)\s*"/, '')
end

