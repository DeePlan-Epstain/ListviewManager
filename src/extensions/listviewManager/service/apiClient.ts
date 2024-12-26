// import axios from 'axios';
// // import config from "../config/config";
// // import { signIn } from '../graph.auth';

// // Create a base Axios instance
// const apiClient = axios.create({
//     baseURL: 'https://epstin100.sharepoint.com/',
//     headers: {
//         "Content-Type": "application/json"
//     },
// });

// const TOKEN_EXPIRATION_TIME = 30 * 60 * 1000;

// // function to check if the token has expired
// const isTokenExpired = (expirationTime: number) => {
//     const currentTime = new Date().getTime();
//     return currentTime > expirationTime;
// };

// apiClient.interceptors.request.use(async (config) => {
//     let token = '';
//     let tokenExpirationTime = 0;

//     const storageToken = localStorage.getItem('token');
//     const storedExpirationTime = localStorage.getItem('tokenExpirationTime');

//     if (storageToken && storedExpirationTime) {
//         token = storageToken;
//         tokenExpirationTime = parseInt(storedExpirationTime, 10);

//         // check if the token is expired
//         if (isTokenExpired(tokenExpirationTime)) token = '';
//     }

//     // if token is missing or expired, get a new one
//     if (!token) {
//         const newToken = await signIn();
//         const newExpirationTime = new Date().getTime() + TOKEN_EXPIRATION_TIME;

//         localStorage.setItem('token', newToken);
//         localStorage.setItem('tokenExpirationTime', newExpirationTime.toString());

//         token = newToken;
//     }

//     if (!token) return Promise.reject(new Error('Invalid token'));

//     config.headers['Authorization'] = `Bearer ${token}`;

//     return config;
// }, (error) => {
//     return Promise.reject(error);
// });

// export default apiClient;
