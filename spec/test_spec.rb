require 'spec_helper'
require 'powerpoint'
require 'powerpoint/util'
include Powerpoint::Util

describe 'Powerpoint parsing a sample PPTX file' do
  before(:all) do

    ##
    # Set up some dummy content based on what comes out of the CMS
    ##
    @what_content ={"global" => {"title" => "Global", "dataType" => "fieldset", "items" => {"subPageNavigationMenu" => {"dataType" => "subPageNavigationMenu"}, "downloadableInput" => {"title" => "Downloadable=> ", "dataType" => "checkBox", "disabled" => true, "items" => ["as PowerPoint", "as PDF"], "value" => ["as PowerPoint", "as PDF"] } } }, "headers" => {"title" => "Header Titles", "dataType" => "fieldset", "items" => {"headerInput1" => {"title" => "Header 1 Title ", "dataType" => "text", "value" => "5 years ago"}, "headerInput2" => {"title" => "Header 2 Title ", "dataType" => "text", "value" => "Now"}, "headerInput3" => {"title" => "Header 3 Title ", "dataType" => "text", "value" => "in 5 years"} } }, "rowsManagerInput" => {"title" => "Row Manager", "dataType" => "manager_items", "disabled" => false, "value" => [{"item" => {"title" => "Row", "dataType" => "fieldset", "items" => {"textInput1" => {"title" => "Column 1 ", "dataType" => "text", "value" => "Social media platforms and review sites extend the range of products and services that consumers can publicly review. In 2012, 340 million tweets were posted per day. "}, "textInput2" => {"title" => "Column 2 ", "dataType" => "text", "value" => "High smartphone ownership allows customers to quickly write or read reviews on the go. Semi-expert consumer voices have arisen in most sectors, whose opinions are given extra weight."}, "textInput3" => {"title" => "Column 3 ", "dataType" => "text", "value" => "An increase in attempts to verify identifies of reviewers in order to rate the trustworthiness of their voices. Fewer sites will allow anonymous reviews to be posted."} } } }, {"item" => {"title" => "Row", "dataType" => "fieldset", "items" => {"textInput1" => {"title" => "Column 1 ", "dataType" => "text", "value" => "Brand interaction with online customer complaints is limited, with customer service efforts focused on in-store feedback and complaints."}, "textInput2" => {"title" => "Column 2 ", "dataType" => "text", "value" => "Brands find new ways of interacting with customers across a range of social media channels, addressing and responding to concerns, complaints and compliments."}, "textInput3" => {"title" => "Column 3 ", "dataType" => "text", "value" => "Inviting and engaging consumers to react, review and critique in advance of product or service launches becomes more common for brands."} } } }, {"item" => {"title" => "Row", "dataType" => "fieldset", "items" => {"textInput1" => {"title" => "Column 1 ", "dataType" => "text", "value" => "CSR efforts become an increasingly common way for brands to enhance their brand and raise their profile."}, "textInput2" => {"title" => "Column 2 ", "dataType" => "text", "value" => "The transparency and lobby-building capacity of social media leads to CSR efforts deemed insincere or ineffectual more easily being called into question."}, "textInput3" => {"title" => "Column 3 ", "dataType" => "text", "value" => "To pre-empt the quick forming of lobbies and factions, brands’ CSR eschew the general and hone in on specifics  -  ensuring the presence of clear goals."} } } }], "itemDefinition" => {"title" => "Row", "dataType" => "fieldset", "items" => {"textInput1" => {"title" => "Column1 ", "dataType" => "text", "value" => "Default body."}, "textInput2" => {"title" => "Column2 ", "dataType" => "text", "value" => "Default body."}, "textInput3" => {"title" => "Column3 ", "dataType" => "text", "value" => "Default body."} } } } }
    @three_col_content = {"global"=>{"title"=>"Global", "dataType"=>"fieldset", "items"=>{"downloadableInput"=>{"title"=>"Downloadable: ", "dataType"=>"checkBox", "disabled"=>true, "items"=>["as PowerPoint", "as PDF"], "value"=>["as PowerPoint", "as PDF"]}}}, "headingInput"=>{"dataType"=>"text", "title"=>"Heading: ", "value"=>"Lorem Ipsum Dolor Sit Amet"}, "column1"=>{"dataType"=>"fieldset", "items"=>{"headingInput"=>{"dataType"=>"text", "title"=>"Heading: ", "value"=>"Lorem Ipsum Dolor Consec"}, "textInput"=>{"dataType"=>"richText", "title"=>"Text: ", "value"=>"<p>Sum expectantes. Ego hodie expectantes. Expectantes, et misit unum de pueris Gus interficere. Et suus vos. Nescio quis, qui est bonus usus liberi ad Isai? Qui nosti ... Quis dimisit filios ad necem ... hmm? Gus! Est, ante me factus singulis decem gradibus. Et nunc ad aliud opus mihi tandem tollendum est puer ille consensus et nunc fugit. Ipse suus obtinuit eam. Non solum autem illa, sed te tractantur in se trahens felis. </p>"}}, "disabled"=>false, "title"=>"Column 1"}, "column2"=>{"dataType"=>"fieldset", "items"=>{"headingInput"=>{"dataType"=>"text", "title"=>"Heading: ", "value"=>"Consectetur Adipiscing"}, "textInput"=>{"dataType"=>"richText", "title"=>"Text: ", "value"=>"<p>Sum expectantes. Ego hodie expectantes. Expectantes, et misit unum de pueris Gus interficere. Et suus vos. Nescio quis, qui est bonus usus liberi ad Isai?</p><p>Qui nosti ... Quis dimisit filios ad necem ... hmm? Gus! Est, ante me factus singulis decem gradibus.</p><p>Et nunc ad aliud opus mihi tandem tollendum est puer ille consensus et nunc fugit. Ipse suus obtinuit eam. Non solum autem illa, sed te tractantur in se trahens felis. </p>"}}, "title"=>"Column 2"}, "column3"=>{"dataType"=>"fieldset", "items"=>{"headingInput"=>{"dataType"=>"text", "title"=>"Heading: ", "value"=>"Pariatur Consectetur"}, "textInput"=>{"dataType"=>"richText", "title"=>"Text: ", "value"=>"<p>Sum expectantes. Ego hodie expectantes. Expectantes, et misit unum de pueris Gus interficere. Et suus vos. Nescio quis, qui est bonus usus liberi ad Isai? Qui nosti ... Quis dimisit filios ad necem ... hmm? Gus! </p><p>Est, ante me factus singulis decem gradibus.</p>"}}, "title"=>"Column 3"}}
    @html = '
      <h1>Test Header</h1><p>
<strong>Bedtime</strong>, and <i>evening</i> time in general, is being re-defined. For a significant number, the hours before sleep can be penetrated by a kind of light work; it is now so easy to curl round a laptop or a tablet and drop your boss an email, while scanning the latest news, while streaming on-demand movies, while online shopping for your mother’s birthday present, and so on.
</p>
<p>
<img height="770" src="samples/images/image4.jpeg" width="1543">
</p>
<p>
Work-life <a href="http://www.spicerack.co.uk">balance</a> is redrawn under wider horizons. This is not just a story of more flexible working hours but a story of work encroaching into those times and places formerly reserved for rest: night time, bedrooms, even holidays. To many millennials, work-life balance is in revolution.
</p>
<p>
More, online media and retail are accessed differently in these arenas. Remote technology for both work and socialising means that consumers are engaging with their devices, and their fellow human beings, in a totally new way.
</p>
<p>
Work-life balance is redrawn under wider horizons. This is not just a story of more flexible working hours but a story of work encroaching into those times and places formerly reserved for rest: night time, bedrooms, even holidays. To many millennials, work-life balance is in revolution.
</p>'
    @sector_content = {"Alcohol"=>{"title"=>"Alcohol", "dataType"=>"fieldset", "items"=>{"impactTextInput"=>{"title"=>"Impact: ", "dataType"=>"richText", "value"=>"<p><a href=\"http://www.foo.com\">Lorem ipsum</a> dolar sit amet consectetur...</p>"}}}, "Beauty and Personal Care"=>{"title"=>"Beauty and Personal Care", "dataType"=>"fieldset", "items"=>{"impactTextInput"=>{"title"=>"Impact: ", "dataType"=>"richText", "value"=>"<p><a href=\"http://www.foo.com\">Lorem ipsum</a>dolar sit amet consectetur...</p>"}}}}

    ##
    # Set up some stuff for testing the sector impact slides
    ##
    @final = Nokogiri::HTML.fragment(@html)
    @header = @final.search('h1').first
    @final.search('h1').first.remove
    @image_paths = []
    @final.search('img').each do |node|
      @image_paths << node.attributes['src']
    end
    sector_image_path = "samples/images/sector_leisure.jpg"
    links = @final.search('a').map { |link| link['href'] }
    sector_impact_links = Nokogiri.parse(@sector_content.first[1]['items'].first[1]['value']).search('a').map{ |link| link['href'] }
    ##
    # Create a new presentation
    ##
    @deck = Powerpoint::Presentation.new

    ##
    # This is where we add slides to the presentation . .testing the templates
    # IF YOU WANT A SINGLE SLIDE TEST use   @deck.add_textual_slide 'test head', ['test body'] and comment out the @embed_deck loop below
    ##


    @deck.add_ff_trend_intro_slide 'Abcdefghijklmnopqrstuvwxyz12345678910112', 'Contactless credit/debit cards, NFC- and web-enabled phones and digital wallets continue to transform the future of payment methods  -  with major implications for the way we will shop and interact with brands in the future. Contactless credit/debit cards, NFC- and web-enabled phones and digital wallets continue to transform the future of payment methods  -  with major implications for the way we will shop and interact with brands in the future.', 'samples/images/image4.jpeg'
    @deck.add_ff_heading_text_slide @header.inner_html, html_to_ooxml(@final.to_s), @image_paths, links
    @deck.add_ff_three_row_text_slide 'What to do', @three_col_content
    @deck.add_ff_what_next_slide 'What will happen next', @what_content
    @deck.add_ff_sector_impact_slide @sector_content.first[1]['title'], @sector_content.first[1]['items'].first[1]['value'], sector_image_path, sector_impact_links
    #@deck.add_ff_associated_content_slide 'Sample Asscociated Content Item', 'Test Associated Content Subtitle', 'samples/images/image4.jpeg', {}, 'sample.pptx'
    @deck.add_ff_trend_outro_slide

    ##
    # These are the embeded prenstatoins I have taken a selection of the ones that have tags, drawiings, charts etc
    #
    #  Loop through each and add the masters and layouts to the main output presentation
    #  The other emedding is taken care of by the lib/powerpoint/ff_embed_slide.rb as it's slide specific behaviour
    ##
    embed_decks = ["samples/pptx/35848.pptx", "samples/pptx/43366.pptx", "samples/pptx/41209.pptx","samples/pptx/TR_EU_Reasons_for_going_on_a_holiday_eb2016_2016.pptx"]

    # comment out this loop for a single slide test
    embed_decks.each do |deck_path|
      @embed_deck = Powerpoint::ReadPresentation.new deck_path

      @master_refs = Hash.new

      @embed_deck.masters.each do |master|
        master_rel_xml = Nokogiri::XML::Document.parse(@embed_deck.files.file.open master.gsub('slideMasters','slideMasters/_rels').gsub('.xml','.xml.rels'))
        theme_path = master_rel_xml.css('Relationship').select{ |node| node['Type'].include? 'theme' }.first['Target']
        theme_xml = Nokogiri::XML::Document.parse(@embed_deck.files.file.open(theme_path.gsub('..','ppt')))
        new_theme_path = @deck.add_theme(theme_xml)[:file_path]
        ##### Loop through rel_xml and pull in media emeds
        embeds = master_rel_xml.css('Relationship').select{ |node|
          node['Target'].include? 'media'
        }
        master_embeds = []
        embeds.each_with_index do |embed,index|
          master_embeds << { id: index, rid: embed['Id'], files: @embed_deck.files, file_path: embed['Target'], content_type: embed['Type'] }
        end
        @master_refs["#{master}".gsub('ppt','..')] = @deck.add_master(Nokogiri::XML::Document.parse(@embed_deck.files.file.open master), new_theme_path, [], master_embeds)
      end

      @layout_refs = Hash.new
      @embed_deck.layouts.each do |layout|
        layout_xml = Nokogiri::XML::Document.parse(@embed_deck.files.file.open layout)
        layout_rel_xml = Nokogiri::XML::Document.parse(@embed_deck.files.file.open layout.gsub('slideLayouts','slideLayouts/_rels').gsub('xml','xml.rels'))
        master = layout_rel_xml.css('Relationship').select{ |node| node['Type'].include? 'slideMaster'}.first['Target']
        layout_rel_xml.css('Relationship').select{ |node|
          if node['Target'].include? 'slideMaster'
            node['Target'] = @master_refs[node['Target']][:file_path]
          end
        }
        @layout_refs["#{layout}".gsub('ppt','..')] = @deck.add_layout(layout_xml, layout_rel_xml, @master_refs[master][:file_path], @embed_deck.files)
      end
      #we've updated the layouts so the actual slide masters need updating with the correct ids
      @deck.update_slide_masters

      @embed_deck.slides.each do |slide|
        @deck.add_ff_embeded_slide slide.title, slide.raw_content, slide.raw_relation_content, slide.images, slide.charts, slide.embeddings, slide.notes, slide.tags, slide.drawings, @master_refs[slide.master], @deck.notes_masters.first, @layout_refs[slide.layout], slide.theme_overrides
      end
    end

    ##
    # Save the presentation
    ##
    @deck.save 'samples/pptx/test-output.pptx'
  end

  it 'Create a PPTX file successfully.' do
    #@deck.should_not be_nil
  end
end
