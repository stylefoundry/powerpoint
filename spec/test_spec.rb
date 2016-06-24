require 'spec_helper'
require 'powerpoint'

describe 'Powerpoint parsing a sample PPTX file' do
  before(:all) do

    @what_content = {"global"=>{"title"=>"Global", "dataType"=>"fieldset", "items"=>{"downloadableInput"=>{"title"=>"Downloadable => ", "dataType"=>"checkBox", "disabled"=>true, "items"=>["as PowerPoint", "as PDF"], "value"=>["as PowerPoint", "as PDF"]}}}, "headers"=>{"title"=>"Header Titles", "dataType"=>"fieldset", "items"=>{"headerInput1"=>{"title"=>"Header 1 Title:", "dataType"=>"text", "value"=>"5 Years ago"}, "headerInput2"=>{"title"=>"Header 2 Title:", "dataType"=>"text", "value"=>"Now"}, "headerInput3"=>{"title"=>"Header 3 Title:", "dataType"=>"text", "value"=>"5 Years time"}}},"rowsManagerInput"=>{"title"=>"Row Manager", "dataType"=>"manager_items", "disabled"=>false, "value"=>[{"title"=>"Row", "dataType"=>"fieldset", "items"=>{"textInput1"=>{"title"=>"Column 1", "dataType"=>"text", "value"=>"1"}, "textInput2"=>{"title"=>"Column 2", "dataType"=>"text", "value"=>"2"}, "textInput3"=>{"title"=>"Column 3", "dataType"=>"text", "value"=>"3"}}}, {"title"=>"Row", "dataType"=>"fieldset", "items"=>{"textInput1"=>{"title"=>"Column 1", "dataType"=>"text", "value"=>"4"}, "textInput2"=>{"title"=>"Column 2", "dataType"=>"text", "value"=>"5"}, "textInput3"=>{"title"=>"Column 3", "dataType"=>"text", "value"=>"6"}}}, {"title"=>"Row", "dataType"=>"fieldset", "items"=>{"textInput1"=>{"title"=>"Column 1", "dataType"=>"text", "value"=>"7"}, "textInput2"=>{"title"=>"Column 2", "dataType"=>"text", "value"=>"8"}, "textInput3"=>{"title"=>"Column 3", "dataType"=>"text", "value"=>"9"}}}]}}
    @three_col_content = {"global"=>{"title"=>"Global", "dataType"=>"fieldset", "items"=>{"downloadableInput"=>{"title"=>"Downloadable: ", "dataType"=>"checkBox", "disabled"=>true, "items"=>["as PowerPoint", "as PDF"], "value"=>["as PowerPoint", "as PDF"]}}}, "headingInput"=>{"dataType"=>"text", "title"=>"Heading: ", "value"=>"Lorem Ipsum Dolor Sit Amet"}, "column1"=>{"dataType"=>"fieldset", "items"=>{"headingInput"=>{"dataType"=>"text", "title"=>"Heading: ", "value"=>"Lorem Ipsum Dolor Consec"}, "textInput"=>{"dataType"=>"richText", "title"=>"Text: ", "value"=>"<p>Sum expectantes. Ego hodie expectantes. Expectantes, et misit unum de pueris Gus interficere. Et suus vos. Nescio quis, qui est bonus usus liberi ad Isai? Qui nosti ... Quis dimisit filios ad necem ... hmm? Gus! Est, ante me factus singulis decem gradibus. Et nunc ad aliud opus mihi tandem tollendum est puer ille consensus et nunc fugit. Ipse suus obtinuit eam. Non solum autem illa, sed te tractantur in se trahens felis. </p>"}}, "disabled"=>false, "title"=>"Column 1"}, "column2"=>{"dataType"=>"fieldset", "items"=>{"headingInput"=>{"dataType"=>"text", "title"=>"Heading: ", "value"=>"Consectetur Adipiscing"}, "textInput"=>{"dataType"=>"richText", "title"=>"Text: ", "value"=>"<p>Sum expectantes. Ego hodie expectantes. Expectantes, et misit unum de pueris Gus interficere. Et suus vos. Nescio quis, qui est bonus usus liberi ad Isai?</p><p>Qui nosti ... Quis dimisit filios ad necem ... hmm? Gus! Est, ante me factus singulis decem gradibus.</p><p>Et nunc ad aliud opus mihi tandem tollendum est puer ille consensus et nunc fugit. Ipse suus obtinuit eam. Non solum autem illa, sed te tractantur in se trahens felis. </p>"}}, "title"=>"Column 2"}, "column3"=>{"dataType"=>"fieldset", "items"=>{"headingInput"=>{"dataType"=>"text", "title"=>"Heading: ", "value"=>"Pariatur Consectetur"}, "textInput"=>{"dataType"=>"richText", "title"=>"Text: ", "value"=>"<p>Sum expectantes. Ego hodie expectantes. Expectantes, et misit unum de pueris Gus interficere. Et suus vos. Nescio quis, qui est bonus usus liberi ad Isai? Qui nosti ... Quis dimisit filios ad necem ... hmm? Gus! </p><p>Est, ante me factus singulis decem gradibus.</p>"}}, "title"=>"Column 3"}}
    @html_content = html_to_ooxml("<p>Test paragraph 1</p><p>Test paragraph 2</p><p><ul><li>list 1</li><li>List 2</li></ul></p> <p>A <b>bold</b> word</p>")
    @sector_content = {"Alcohol"=>{"title"=>"Alcohol", "dataType"=>"fieldset", "items"=>{"impactTextInput"=>{"title"=>"Impact: ", "dataType"=>"richText", "value"=>"<p>Lorem ipsum dolar sit amet consectetur...</p>"}}}, "Beauty and Personal Care"=>{"title"=>"Beauty and Personal Care", "dataType"=>"fieldset", "items"=>{"impactTextInput"=>{"title"=>"Impact: ", "dataType"=>"richText", "value"=>"<p>Lorem ipsum dolar sit amet consectetur...</p>"}}}}
    @deck = Powerpoint::Presentation.new
    @deck.add_intro 'Test Presentation', 'created automagically'
    @deck.add_ff_trend_intro_slide 'Cashless Society', 'Contactless credit/debit cards, NFC- and web-enabled phones and digital wallets continue to transform the future of payment methods  -  with major implications for the way we will shop and interact with brands in the future.', 'samples/images/image4.jpeg'
    @deck.add_ff_trend_what_next_slide 'What will happen next', @what_content
    @deck.add_ff_trend_heading_text_slide 'Test Header', @html_content
    @deck.add_ff_trend_three_row_text_slide 'What to do', @three_col_content
    @deck.add_ff_trend_sector_impact_slide 'Sector Impact', @sector_content
    @deck.save 'samples/pptx/sample.pptx' # Examine the PPTX file
  end

  it 'Create a PPTX file successfully.' do
    #@deck.should_not be_nil
  end
end
