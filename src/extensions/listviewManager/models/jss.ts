import { create } from "jss";
import rtl from 'jss-rtl';
import { jssPreset } from '@material-ui/core/styles';

export const jss = create({
    plugins: [...jssPreset().plugins, rtl()]
});