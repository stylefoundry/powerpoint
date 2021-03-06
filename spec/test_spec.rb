require 'spec_helper'
require 'powerpoint'
require 'powerpoint/util'
include Powerpoint::Util

describe 'Powerpoint parsing a sample PPTX file' do
  before(:all) do

    @what_content ={"global" => {"title" => "Global", "dataType" => "fieldset", "items" => {"subPageNavigationMenu" => {"dataType" => "subPageNavigationMenu"}, "downloadableInput" => {"title" => "Downloadable=> ", "dataType" => "checkBox", "disabled" => true, "items" => ["as PowerPoint", "as PDF"], "value" => ["as PowerPoint", "as PDF"] } } }, "headers" => {"title" => "Header Titles", "dataType" => "fieldset", "items" => {"headerInput1" => {"title" => "Header 1 Title ", "dataType" => "text", "value" => "5 years ago"}, "headerInput2" => {"title" => "Header 2 Title ", "dataType" => "text", "value" => "Now"}, "headerInput3" => {"title" => "Header 3 Title ", "dataType" => "text", "value" => "in 5 years"} } }, "rowsManagerInput" => {"title" => "Row Manager", "dataType" => "manager_items", "disabled" => false, "value" => [{"item" => {"title" => "Row", "dataType" => "fieldset", "items" => {"textInput1" => {"title" => "Column 1 ", "dataType" => "text", "value" => "Social media platforms and review sites extend the range of products and services that consumers can publicly review. In 2012, 340 million tweets were posted per day. "}, "textInput2" => {"title" => "Column 2 ", "dataType" => "text", "value" => "High smartphone ownership allows customers to quickly write or read reviews on the go. Semi-expert consumer voices have arisen in most sectors, whose opinions are given extra weight."}, "textInput3" => {"title" => "Column 3 ", "dataType" => "text", "value" => "An increase in attempts to verify identifies of reviewers in order to rate the trustworthiness of their voices. Fewer sites will allow anonymous reviews to be posted."} } } }, {"item" => {"title" => "Row", "dataType" => "fieldset", "items" => {"textInput1" => {"title" => "Column 1 ", "dataType" => "text", "value" => "Brand interaction with online customer complaints is limited, with customer service efforts focused on in-store feedback and complaints."}, "textInput2" => {"title" => "Column 2 ", "dataType" => "text", "value" => "Brands find new ways of interacting with customers across a range of social media channels, addressing and responding to concerns, complaints and compliments."}, "textInput3" => {"title" => "Column 3 ", "dataType" => "text", "value" => "Inviting and engaging consumers to react, review and critique in advance of product or service launches becomes more common for brands."} } } }, {"item" => {"title" => "Row", "dataType" => "fieldset", "items" => {"textInput1" => {"title" => "Column 1 ", "dataType" => "text", "value" => "CSR efforts become an increasingly common way for brands to enhance their brand and raise their profile."}, "textInput2" => {"title" => "Column 2 ", "dataType" => "text", "value" => "The transparency and lobby-building capacity of social media leads to CSR efforts deemed insincere or ineffectual more easily being called into question."}, "textInput3" => {"title" => "Column 3 ", "dataType" => "text", "value" => "To pre-empt the quick forming of lobbies and factions, brands’ CSR eschew the general and hone in on specifics  -  ensuring the presence of clear goals."} } } }], "itemDefinition" => {"title" => "Row", "dataType" => "fieldset", "items" => {"textInput1" => {"title" => "Column1 ", "dataType" => "text", "value" => "Default body."}, "textInput2" => {"title" => "Column2 ", "dataType" => "text", "value" => "Default body."}, "textInput3" => {"title" => "Column3 ", "dataType" => "text", "value" => "Default body."} } } } }
    @three_col_content = {"global"=>{"title"=>"Global", "dataType"=>"fieldset", "items"=>{"downloadableInput"=>{"title"=>"Downloadable: ", "dataType"=>"checkBox", "disabled"=>true, "items"=>["as PowerPoint", "as PDF"], "value"=>["as PowerPoint", "as PDF"]}}}, "headingInput"=>{"dataType"=>"text", "title"=>"Heading: ", "value"=>"Lorem Ipsum Dolor Sit Amet"}, "column1"=>{"dataType"=>"fieldset", "items"=>{"headingInput"=>{"dataType"=>"text", "title"=>"Heading: ", "value"=>"Lorem Ipsum Dolor Consec"}, "textInput"=>{"dataType"=>"richText", "title"=>"Text: ", "value"=>"<p>Sum expectantes. Ego hodie expectantes. Expectantes, et misit unum de pueris Gus interficere. Et suus vos. Nescio quis, qui est bonus usus liberi ad Isai? Qui nosti ... Quis dimisit filios ad necem ... hmm? Gus! Est, ante me factus singulis decem gradibus. Et nunc ad aliud opus mihi tandem tollendum est puer ille consensus et nunc fugit. Ipse suus obtinuit eam. Non solum autem illa, sed te tractantur in se trahens felis. </p>"}}, "disabled"=>false, "title"=>"Column 1"}, "column2"=>{"dataType"=>"fieldset", "items"=>{"headingInput"=>{"dataType"=>"text", "title"=>"Heading: ", "value"=>"Consectetur Adipiscing"}, "textInput"=>{"dataType"=>"richText", "title"=>"Text: ", "value"=>"<p>Sum expectantes. Ego hodie expectantes. Expectantes, et misit unum de pueris Gus interficere. Et suus vos. Nescio quis, qui est bonus usus liberi ad Isai?</p><p>Qui nosti ... Quis dimisit filios ad necem ... hmm? Gus! Est, ante me factus singulis decem gradibus.</p><p>Et nunc ad aliud opus mihi tandem tollendum est puer ille consensus et nunc fugit. Ipse suus obtinuit eam. Non solum autem illa, sed te tractantur in se trahens felis. </p>"}}, "title"=>"Column 2"}, "column3"=>{"dataType"=>"fieldset", "items"=>{"headingInput"=>{"dataType"=>"text", "title"=>"Heading: ", "value"=>"Pariatur Consectetur"}, "textInput"=>{"dataType"=>"richText", "title"=>"Text: ", "value"=>"<p>Sum expectantes. Ego hodie expectantes. Expectantes, et misit unum de pueris Gus interficere. Et suus vos. Nescio quis, qui est bonus usus liberi ad Isai? Qui nosti ... Quis dimisit filios ad necem ... hmm? Gus! </p><p>Est, ante me factus singulis decem gradibus.</p>"}}, "title"=>"Column 3"}}
    @html_content = html_to_ooxml('<p>
Bedtime, and evening time in general, is being re-defined. For a significant number, the hours before sleep can be penetrated by a kind of light work; it is now so easy to curl round a laptop or a tablet and drop your boss an email, while scanning the latest news, while streaming on-demand movies, while online shopping for your mother’s birthday present, and so on.
</p>
<p>
The consumer skips between sites and roles, merging personal and private life in these hours, in a state of fused brainstorming and multitasking.
</p>
<p>
Work-life balance is redrawn under wider horizons. This is not just a story of more flexible working hours but a story of work encroaching into those times and places formerly reserved for rest: night time, bedrooms, even holidays. To many millennials, work-life balance is in revolution.
</p>
<p>
More, online media and retail are accessed differently in these arenas. Remote technology for both work and socialising means that consumers are engaging with their devices, and their fellow human beings, in a totally new way.
</p>
<h1>
Global differences
</h1>
<img alt data-ce-max-width height="770" src="https://nvision-public.s3.amazonaws.com/uploads/asset/file/20/Map_-_China_India_Mexico_SKorea_US_UK_Aus.png" width="1543">
<p>
Usage of tech in bed before going to sleep is pronounced globally, with at least a third doing so across all markets surveyed. However, we see particularly high rates of portable device usage in South Korea, where smartphone ownership is high, and across emerging markets, where cheaper, tablet or smartphone ownership will sometimes replace laptop ownership. Watching TV on demand is still most commonly carried out via laptop, suggesting that late-night smartphone and tablet use is more likely to be focused on other activities, such as social media, email and messaging. Twitter activity in the US, UK, Australia and India peaks at bedtime, around 10pm. Morning bedtime Twitter activity is also on the rise.
</p>
<p>
While checking personal emails in bed is, on the whole, more popular than checking work emails. In many emerging markets, such as India, Mexico and China, the difference is minimal suggesting a more entrenched 24/7 switched on culture. In Asia, this trend is further boosted by late night shop openings and typically later working hours than in Western markets.
</p>')
    @sector_content = {"Alcohol"=>{"title"=>"Alcohol", "dataType"=>"fieldset", "items"=>{"impactTextInput"=>{"title"=>"Impact: ", "dataType"=>"richText", "value"=>"<p>Lorem ipsum dolar sit amet consectetur...</p>"}}}, "Beauty and Personal Care"=>{"title"=>"Beauty and Personal Care", "dataType"=>"fieldset", "items"=>{"impactTextInput"=>{"title"=>"Impact: ", "dataType"=>"richText", "value"=>"<p>Lorem ipsum dolar sit amet consectetur...</p>"}}}}

    @deck = Powerpoint::Presentation.new
    @deck.add_ff_trend_intro_slide 'Cashless Society', 'Contactless credit/debit cards, NFC- and web-enabled phones and digital wallets continue to transform the future of payment methods  -  with major implications for the way we will shop and interact with brands in the future.', 'samples/images/image4.jpeg'
    @deck.add_ff_what_next_slide 'What will happen next', @what_content
    @deck.add_ff_heading_text_slide 'Test Header', @html_content
    @deck.add_ff_three_row_text_slide 'What to do', @three_col_content
    @deck.add_ff_sector_impact_slide 'Sector Impact', @sector_content
    @deck.add_ff_associated_content_slide 'Sample Asscociated Content Item', 'Test Associated Content Subtitle', 'samples/images/image4.jpeg', {}, 'sample.pptx'

    @deck.save 'samples/pptx/test-output.pptx'
  end

  it 'Create a PPTX file successfully.' do
    #@deck.should_not be_nil
  end
end
