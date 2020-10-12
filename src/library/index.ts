import utils from './utils';

export default function echo(data : any, err : Error) {
  return new Promise((resolve, reject) => {
    if (err) {
      return reject(err);
    }
    return resolve(utils(data));
  });
}
