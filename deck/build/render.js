const React=require('react');const {renderToStaticMarkup}=require('react-dom/server');const sharp=require('sharp');
const fa=require('react-icons/fa'); const tb=require('react-icons/tb'); const md=require('react-icons/md');
const icons={coins:fa.FaCoins,speed:md.MdSpeed,clock:fa.FaRegClock,handshake:fa.FaHandshake,robot:fa.FaRobot,search:fa.FaSearch,loop:md.MdLoop,sum:tb.TbSum,layers:fa.FaLayerGroup,filter:fa.FaFilter,sliders:fa.FaSlidersH,play:fa.FaPlay,bolt:fa.FaBolt,database:fa.FaDatabase,list:fa.FaListUl,check:fa.FaCheck,times:fa.FaTimes,code:fa.FaCode,chart:fa.FaChartBar,link:fa.FaExternalLinkAlt,table:fa.FaTable,lock:fa.FaLock,arrow:fa.FaArrowRight,stopwatch:fa.FaStopwatch,columns:fa.FaColumns,sitemap:fa.FaSitemap,hashtag:fa.FaHashtag,random:fa.FaRandom,cube:fa.FaCube,flag:fa.FaFlagCheckered,github:fa.FaGithub,book:fa.FaBook,excel:fa.FaFileExcel,mobile:fa.FaMobileAlt,stream:fa.FaStream,key:fa.FaKey,exchange:fa.FaExchangeAlt,warning:fa.FaExclamationTriangle,eye:fa.FaEye,redo:fa.FaRedo,pages:fa.FaCopy};
const colors={w:'#FFFFFF',n:'#040E3D',b:'#1E6FFF'};
(async()=>{for(const [k,I] of Object.entries(icons)){ if(!I){console.log('MISSING',k);continue;}
 for(const [ck,c] of Object.entries(colors)){const svg=renderToStaticMarkup(React.createElement(I,{size:256,color:c}));
 await sharp(Buffer.from(svg)).resize(256,256,{fit:'contain',background:{r:0,g:0,b:0,alpha:0}}).png().toFile(`${k}_${ck}.png`);}}
 console.log('done')})();
