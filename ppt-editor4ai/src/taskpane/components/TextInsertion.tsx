import * as React from "react";
import { useState } from "react";
import { Button, Field, Textarea, Input, tokens, makeStyles } from "@fluentui/react-components";

/* global HTMLTextAreaElement, HTMLInputElement */

interface TextInsertionProps {
  insertText: (text: string, left?: number, top?: number) => void;
}

const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  textAreaField: {
    marginLeft: "20px",
    marginTop: "30px",
    marginBottom: "20px",
    marginRight: "20px",
    maxWidth: "50%",
  },
  positionContainer: {
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    marginLeft: "20px",
    marginRight: "20px",
    marginBottom: "10px",
    maxWidth: "50%",
  },
  positionField: {
    width: "100%",
  },
  hint: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginLeft: "20px",
    marginRight: "20px",
    marginBottom: "15px",
    maxWidth: "50%",
    textAlign: "center",
  },
});

const TextInsertion: React.FC<TextInsertionProps> = (props: TextInsertionProps) => {
  const [text, setText] = useState<string>("Some text.");
  const [left, setLeft] = useState<string>("");
  const [top, setTop] = useState<string>("");

  const handleTextInsertion = async () => {
    // 解析位置参数 / Parse position parameters
    const leftValue = left.trim() === "" ? undefined : parseFloat(left);
    const topValue = top.trim() === "" ? undefined : parseFloat(top);
    
    await props.insertText(text, leftValue, topValue);
  };

  const handleTextChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(event.target.value);
  };

  const handleLeftChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    setLeft(event.target.value);
  };

  const handleTopChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    setTop(event.target.value);
  };

  const styles = useStyles();

  return (
    <div className={styles.textPromptAndInsertion}>
      <Field className={styles.textAreaField} size="large" label="输入待插入文本">
        <Textarea size="large" value={text} onChange={handleTextChange} />
      </Field>
      
      <div className={styles.positionContainer}>
        <Field className={styles.positionField} label="X 坐标 (可选)">
          <Input 
            type="number" 
            value={left} 
            onChange={handleLeftChange}
            placeholder="留空使用默认"
          />
        </Field>
        <Field className={styles.positionField} label="Y 坐标 (可选)">
          <Input 
            type="number" 
            value={top} 
            onChange={handleTopChange}
            placeholder="留空使用默认"
          />
        </Field>
      </div>
      
      <div className={styles.hint}>
        💡 位置范围提示: <br />
        标准 16:9 幻灯片尺寸约为 720×540 磅 (points)<br />
        X 范围: 0-720, Y 范围: 0-540
      </div>
      
      <Field className={styles.instructions}>点击正文按钮插入.</Field>
      <Button appearance="primary" disabled={false} size="large" onClick={handleTextInsertion}>
        确认插入
      </Button>
    </div>
  );
};

export default TextInsertion;
