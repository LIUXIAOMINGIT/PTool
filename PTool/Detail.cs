using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Misc = ComunicationProtocol.Misc;

namespace PTool
{
    public partial class Detail : UserControl
    {
        private float m_P0 = 0;
        private List<PressureCalibrationParameter> m_CaliParameters = new List<PressureCalibrationParameter>();
        private PumpID m_LocalPid = PumpID.GrasebyC6;//默认显示的是C6
        private int  m_Channel = 1;//1号通道，默认值

        public float P0
        {
            set { m_P0 = value; }
            get { return m_P0; }
        }

        /// <summary>
        /// 产品ID
        /// </summary>
        public PumpID Pid
        {
            set { m_LocalPid = value; }
            get { return m_LocalPid; }
        }

        /// <summary>
        /// 第几道泵的详细信息
        /// </summary>
        public int Channel
        {
            set { m_Channel = value; }
            get { return m_Channel; }
        }

        [Bindable(false), Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        /// <summary>
        /// 
        /// </summary>
        public List<PressureCalibrationParameter> CaliParameters
        {
            set { m_CaliParameters = value; }
            get { return m_CaliParameters; }
        }


        public Detail()
        {
            InitializeComponent();
        }

        public void ClearLabelValue()
        {
            lbL10Right.Text = "";
            lbC10Right.Text = "";
            lbH10Right.Text = "";
            lbL10Left.Text = "";
            lbC10Left.Text = "";
            lbH10Left.Text = "";
            lbL20Right.Text = "";
            lbC20Right.Text = "";
            lbH20Right.Text = "";
            lbL20Left.Text = "";
            lbC20Left.Text = "";
            lbH20Left.Text = "";
            lbL30Right.Text = "";
            lbC30Right.Text = "";
            lbH30Right.Text = "";
            lbL30Left.Text = "";
            lbC30Left.Text = "";
            lbH30Left.Text = "";
            lbL50Right.Text = "";
            lbC50Right.Text = "";
            lbH50Right.Text = "";
            lbL50Left.Text = "";
            lbC50Left.Text = "";
            lbH50Left.Text = "";
        }

        private void picClose_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void Detail_VisibleChanged(object sender, EventArgs e)
        {
            PumpID pid = PumpID.None;
            if (this.m_Channel == 1)
            {
                switch (m_LocalPid)
                {
                    case PumpID.GrasebyF6_2:
                        pid = PumpID.GrasebyF6;
                        break;
                    case PumpID.WZS50F6_2:
                        pid = PumpID.WZS50F6;
                        break;
                    default:
                        pid = m_LocalPid;
                        break;
                }
            }
            else
            {
                pid = m_LocalPid;
            }

            if(this.Visible)
            {
                lbP0.Text = "P0="+this.m_P0.ToString("F0");
                PressureCalibrationParameter para = null;
                para = m_CaliParameters.Find((x) => { return x.m_SyringeSize == 10; });
                if (para != null)
                {
                    lbL10Right.Text = (para.m_PressureL * 100).ToString("F0");
                    lbC10Right.Text = (para.m_PressureC * 100).ToString("F0");
                    lbH10Right.Text = (para.m_PressureH * 100).ToString("F0");
                    lbL10Left.Text = PressureManager.Instance().GetMidBySizeLevel(pid, 10, Misc.OcclusionLevel.L).ToString("F2");
                    lbC10Left.Text = PressureManager.Instance().GetMidBySizeLevel(pid, 10, Misc.OcclusionLevel.C).ToString("F2");
                    lbH10Left.Text = PressureManager.Instance().GetMidBySizeLevel(pid, 10, Misc.OcclusionLevel.H).ToString("F2");
                }
                para = m_CaliParameters.Find((x) => { return x.m_SyringeSize == 20; });
                if (para != null)
                {
                    lbL20Right.Text = (para.m_PressureL * 100).ToString("F0");
                    lbC20Right.Text = (para.m_PressureC * 100).ToString("F0");
                    lbH20Right.Text = (para.m_PressureH * 100).ToString("F0");
                    lbL20Left.Text = PressureManager.Instance().GetMidBySizeLevel(pid, 20, Misc.OcclusionLevel.L).ToString("F2");
                    lbC20Left.Text = PressureManager.Instance().GetMidBySizeLevel(pid, 20, Misc.OcclusionLevel.C).ToString("F2");
                    lbH20Left.Text = PressureManager.Instance().GetMidBySizeLevel(pid, 20, Misc.OcclusionLevel.H).ToString("F2");
                }

                para = m_CaliParameters.Find((x) => { return x.m_SyringeSize == 30; });
                if (para != null)
                {
                    lbL30Right.Text = (para.m_PressureL * 100).ToString("F0");
                    lbC30Right.Text = (para.m_PressureC * 100).ToString("F0");
                    lbH30Right.Text = (para.m_PressureH * 100).ToString("F0");
                    lbL30Left.Text = PressureManager.Instance().GetMidBySizeLevel(pid,30, Misc.OcclusionLevel.L).ToString("F2");
                    lbC30Left.Text = PressureManager.Instance().GetMidBySizeLevel(pid,30, Misc.OcclusionLevel.C).ToString("F2");
                    lbH30Left.Text = PressureManager.Instance().GetMidBySizeLevel(pid,30, Misc.OcclusionLevel.H).ToString("F2");
                }

                para = m_CaliParameters.Find((x) => { return x.m_SyringeSize == 50; });
                if (para != null)
                {
                    lbL50Right.Text = (para.m_PressureL * 100).ToString("F0");
                    lbC50Right.Text = (para.m_PressureC * 100).ToString("F0");
                    lbH50Right.Text = (para.m_PressureH * 100).ToString("F0");
                    lbL50Left.Text = PressureManager.Instance().GetMidBySizeLevel(pid, 50, Misc.OcclusionLevel.L).ToString("F2");
                    lbC50Left.Text = PressureManager.Instance().GetMidBySizeLevel(pid, 50, Misc.OcclusionLevel.C).ToString("F2");
                    lbH50Left.Text = PressureManager.Instance().GetMidBySizeLevel(pid, 50, Misc.OcclusionLevel.H).ToString("F2");
 
                }
            }
        }

       
    }
}
