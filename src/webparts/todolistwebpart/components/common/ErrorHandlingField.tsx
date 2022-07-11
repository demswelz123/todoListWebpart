import * as React from "react";
import { Label } from "office-ui-fabric-react";
export type ErrorHandlingFieldProps = {
  parentClass?: string;
  labelClass?: string;
  label: string;
  isRequired: boolean;
  errorMessage: string;
  key?: string;
  children: any;
  styleOverride?: any;
  overrideTrigger?: any;
};
const validateVWTrigger = (condition: {gt?, ge?, eq?, lt?, le?} = {} ) => {
    const a = window.innerWidth;
    const {gt, ge, eq, lt, le} = condition;
    let check = 0;
    let pass = 0;
    if (gt) {
        check++;
        a > gt ? pass++ : pass--;
    }
    if (ge) {
        check++;
        a >= ge ? pass++ : pass--;
    }
    if (eq) {
        check++;
        a == eq ? pass++ : pass--;
    }
    if (lt) {
        check++;
        a > lt ? pass++ : pass--;
    }
    if (le) {
        check++;
        a > le ? pass++ : pass--;
    }
    // console.log('check', check);
    // console.log('pass', pass);
    return check === pass ? true : false;
};
const ErrorHandlingField = (props: ErrorHandlingFieldProps) => {
    const ov = validateVWTrigger(props.overrideTrigger);
    return (
        <div className={props.parentClass} style={ov ? props.styleOverride : {}}>
            <Label required={props.isRequired} className={props.labelClass}>
                {props.label}
            </Label>
            <div style={props.errorMessage && { border: "1px solid #a80000" }}>
                {props.children}
            </div>
            {props.errorMessage && (
                <div style={{ color: "#a80000", fontSize: "12px", fontWeight: 400 }}>
                    {props.errorMessage}
                </div>
            )}
        </div>
    );
};
export default ErrorHandlingField;