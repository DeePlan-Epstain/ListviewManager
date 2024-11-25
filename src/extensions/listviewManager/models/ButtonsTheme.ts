import { createTheme } from '@material-ui/core/styles';
export const ButtonsTheme = createTheme({
    direction: 'rtl', // Both here and <body dir="rtl">
    palette: {
        primary: {
            main: '#6DB97E',
            contrastText: "#fff"
        },
        secondary: {
            main: '#C7495D'
        }
    },
});