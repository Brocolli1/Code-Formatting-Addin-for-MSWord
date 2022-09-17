import * as React from 'react';
import { PrimaryButton, IButtonProps } from '@fluentui/react/lib/Button';
import { Label } from '@fluentui/react/lib/Label';
import { string } from 'prop-types';
import { ContextReplacementPlugin, LoaderContext } from 'webpack';

var level = 0;

export class ButtonPrimaryExample extends React.Component<IButtonProps, {}> {
  public constructor(props) {
    super(props);
  }

  insertText = async () => {
    await Word.run(async (context) => {

      var keywords = {
        "if" : "magenta",
        "else" : "magenta",
        "import" : "magenta",
        "for" : "magenta",
        "return" : "magenta",
        "var" : "blue",
        "int" : "blue"
      };
 
      const paragraphs = context.document.getSelection().paragraphs;
      var words;

      var color:string = "black";
      
      level = 0;

      paragraphs.load();
      await context.sync();
      for (let i = 0; i < paragraphs.items.length; i++) {
        
        var paragraph = paragraphs.items[i];
        await this.paragraphBreaker(paragraph,keywords);
      }
    });
  }

  private async paragraphBreaker(paragraph, keywords){
    paragraph.load("text");
    var context = paragraph.context;
    var words;
    var color = "black";
    paragraph.firstLineIndent = level * 10; 
    await context.sync();
      words = paragraph.split(["{", "}", "(", ")", ";", "\t", " ", "\n"], false, true);
      words.load("text");
      await context.sync();

      for (let j = 0; j < words.items.length; j++) {
        
        var word = words.items[j].text;

        if(word.includes("{"))
        {
          level += 1; 
          if(j != words.items.length-1){
            //words.items[j].insertText("\n", "End");
            words.items[j].insertBreak("Line", "After");
            await context.sync();
            //await this.paragraphBreaker(paragraph.getNext(), keywords);
          }  
        }
        if(word.includes("}"))
        {
          await context.sync();
          if(j == words.items.length-1){
            await context.sync();
          } 
          else{
            words.items[j].insertBreak("Line", "After");
            await context.sync();
          }
        }
        if(word.includes(";"))
        {
          if(j != words.items.length-1){
            words.items[j+1].insertBreak("Line", "Before");
            await context.sync();
            //await this.paragraphBreaker(paragraph.getNext(), keywords);
          }
        }
        if(word.includes("\"") || word.includes("\'"))
        {
          if(color == "red"){
            color = "black";
            await this.colorChange("black", "red", words.items[j], "\"", false);
          }
          else {
            color = "red";
            await this.colorChange("red", "black", words.items[j], "\"", true);
          }
        }
        if(word.includes("/*"))
        {
          await this.colorChange("green", color, words.items[j], "/*", true);
          color = "green";
        }
        if(word.includes("*/"))
        {
          if(color == "green")
          {
            await this.colorChange("black", "green", words.items[j], "*/", false);
            color = "black";
          }
          
        }
        else {
          words.items[j].font.color = color;
          await context.sync();
        }

        await this.matchKeyword(words.items[j], keywords);
        await context.sync();
      }
  }

  private async colorChange(newColor:string, oldColor:string, range:Word.Range, delimeter:string, inclusive:boolean)
  {
    var workingRange = range.split([delimeter], true, inclusive);
    workingRange.load("text");

    if(inclusive) {
      range.font.color = newColor; 
      await range.context.sync();
    }
    else {
      range.font.color = oldColor; 
      await range.context.sync();
    }

    if(inclusive && workingRange.items.length > 1)
    {
      workingRange.items[0].font.color = oldColor; 
      await workingRange.context.sync();
    }
    else 
    {  
      if (workingRange.items.length > 1){
        workingRange.items[workingRange.items.length - 1].font.color = newColor;
        await range.context.sync();
      }
    }
    
  }

  private async writeLine(text:String, level:Number, context)
  { 
    if(text.trim().length != 0) {
      context.document.body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    }
    return level
  }

  private async matchKeyword(range:Word.Range, keywords){
    //await this.writeLine(range.text, 0, range.context);
    
    var workingRange = range.split(['{'], true, true);
    workingRange.load("text");

    var word = range.text.replace(/[^a-zA-Z ]/g, "");
    
    for(var key in keywords)
    {
      if (word == key)
      {
        range.font.color = keywords[key];
        await range.context.sync();
      }
    }
  }

  public render() {
    let { disabled } = this.props;
    return (
      <div className='ms-BasicButtonsExample'>
        <Label>Click the button to format selected code</Label>
        <PrimaryButton
          data-automation-id='test'
          disabled={ disabled }
          text='Format code'
          onClick={ this.insertText } />
      </div>
    );
  }
}