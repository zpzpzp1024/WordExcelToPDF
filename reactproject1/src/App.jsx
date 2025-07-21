import {useState} from 'react'
import './App.css'
import {InboxOutlined} from '@ant-design/icons';
import {Button, message, Modal, Spin, Upload, Alert} from 'antd';
import React from 'react';

const {Dragger} = Upload;
const props = {
    name: 'file',
    accept: '.xls,.xlsx,.xlsm,.doc,.docx',
    multiple: true,
    showUploadList: false,
    onChange(info) {
        console.log(info);

        //     const { status } = info.file;
        //     if (status !== 'uploading') {
        //         console.log(info.file, info.fileList);
        //     }
        //     if (status === 'done') {
        //         message.success(`${info.file.name} file uploaded successfully.`);
        //     } else if (status === 'error') {
        //         message.error(`${info.file.name} file upload failed.`);
        //     }
        // },
        // onDrop(e) {
        //     console.log('Dropped files', e.dataTransfer.files);
    },

};

function App() {
    const [loading, setLoading] = React.useState(false);
    const [totalFileCount, setTotalFileCount] = React.useState(0);
    const handleFileSelect = async (file, fileList) => {
        try {
            const reader = new FileReader();
            reader.onload = (e) => {
                const base64Data = e.target.result.split(',')[1]; // 移除data:前缀

                if (window.chrome && window.chrome.webview) {
                    window.chrome.webview.hostObjects.FileHelper.ReceiveFileAsBase64(base64Data, file.name);
                }
            };
            reader.readAsDataURL(file);
            setTotalFileCount(fileList.length);
        } catch (error) {
            console.error('文件处理错误:', error);
        }
    };

    const info = () => {
        Modal.info({
            title: 'This is a notification message',
            content: (
                <div>
                    <p>some messages...some messages...</p>
                    <p>some messages...some messages...</p>
                </div>
            ),
            onOk() {
            },
        });
    };
    const success = () => {
        Modal.success({
            content: 'some messages...some messages...',
        });
    };
    const error = () => {
        Modal.error({
            title: 'This is an error message',
            content: 'some messages...some messages...',
        });
    };
    const warning = () => {
        Modal.warning({
            title: 'This is a warning message',
            content: 'some messages...some messages...',
        });
    };

    async function ConvertAndMergeToPDF() {
        if(totalFileCount === 0){
            Modal.error({
                content: '请先上传文件',
            });
            return;
        }
        setLoading(true);
        if (window.chrome && window.chrome.webview) {
            var ret = await window.chrome.webview.hostObjects.FileHelper.ConvertAndMergeToPDF();
            if (ret) {
                setLoading(false);
                Modal.success({
                    content: '转换成功，文件保存在桌面上',
                });
                setTotalFileCount(0);
            } else {
                setLoading(false);
                Modal.error({
                    content: '转换失败',
                });
            }

        } else {
            setLoading(false);
            Modal.error({
                content: '转换失败',
            });
        }
    }

    return (


        <Spin spinning={loading}>

            <div style={{width: '100%', height: 'calc(100vh - 50px)'}}>

                <Dragger {...props} beforeUpload={(file, fileList) => {
                    handleFileSelect(file, fileList);
                    return false;
                }}>
                    <p className="ant-upload-drag-icon">
                        <InboxOutlined/>
                    </p>
                    <p className="ant-upload-text">点击或者拖动到这里开始转换</p>
                    <p className="ant-upload-hint">
                        支持单个和批量导入，格式为xls、xlsx、xlsm、doc、docx
                    </p>
                    <br/>
                    {totalFileCount > 0 && <Alert message={`已上传了 ${totalFileCount} 个文件`} type="info" showIcon/>}
                </Dragger>
            </div>
            <Button style={{height: '50px', fontSize: '18px'}} onClick={() => ConvertAndMergeToPDF()} block
                    type="primary">开始转换</Button>
        </Spin>


    )
}

export default App
